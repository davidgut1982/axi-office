/**
 * Why: Provides curated 3-4 field TOON schema helpers so every @axi-office CLI tool
 * can render consistent, token-efficient TOON output without duplicating field extraction
 * logic across packages.
 * What: Exports field schema constructors and extract/render utilities that produce
 * structured TOON-encoded output via @toon-format/toon's encode().
 * Test: Construct a schema with field/pluck/relativeTime, call extract() with a mock
 * item, assert the returned object has the expected keys and values; call renderList()
 * with a single-element array and assert the output string contains the label and value.
 */
import { encode } from "@toon-format/toon";

// ---------------------------------------------------------------------------
// Schema node types
// ---------------------------------------------------------------------------

export interface FieldDef {
	type: "field";
	key: string;
	as: string;
}

export interface PluckDef {
	type: "pluck";
	key: string;
	subkey: string;
	as: string;
}

export interface JoinArrayDef {
	type: "joinArray";
	key: string;
	subkey: string;
	as: string;
	empty: string;
}

export interface RelativeTimeDef {
	type: "relativeTime";
	key: string;
	as: string;
}

export interface BoolYesNoDef {
	type: "boolYesNo";
	key: string;
	as: string;
}

export interface MapEnumDef {
	type: "mapEnum";
	key: string;
	map: Record<string, string>;
	fallback: string;
	as: string;
}

export interface LowerDef {
	type: "lower";
	key: string;
	as: string;
}

export interface CustomDef {
	type: "custom";
	as: string;
	fn: (item: Record<string, unknown>) => unknown;
}

export type SchemaDef =
	| FieldDef
	| PluckDef
	| JoinArrayDef
	| RelativeTimeDef
	| BoolYesNoDef
	| MapEnumDef
	| LowerDef
	| CustomDef;

// ---------------------------------------------------------------------------
// Schema constructors
// ---------------------------------------------------------------------------

/** Extract a top-level field from an item. */
export function field(key: string, as?: string): FieldDef {
	return { type: "field", key, as: as ?? key };
}

/** Extract a nested field (item[key][subkey]). */
export function pluck(key: string, subkey: string, as?: string): PluckDef {
	return { type: "pluck", key, subkey, as: as ?? subkey };
}

/** Join an array of strings or objects (using subkey) into a comma-separated value. */
export function joinArray(key: string, subkey: string, as?: string, empty = "none"): JoinArrayDef {
	return { type: "joinArray", key, subkey, as: as ?? key, empty };
}

/** Format an ISO timestamp as a human-readable relative time (e.g. "3h ago"). */
export function relativeTime(key: string, as?: string): RelativeTimeDef {
	return { type: "relativeTime", key, as: as ?? key };
}

/** Render a boolean field as "yes"/"no". */
export function boolYesNo(key: string, as?: string): BoolYesNoDef {
	return { type: "boolYesNo", key, as: as ?? key };
}

/** Map an enum string value through a lookup table; use fallback when not found. */
export function mapEnum(
	key: string,
	map: Record<string, string>,
	fallback: string,
	as?: string
): MapEnumDef {
	return { type: "mapEnum", key, map, fallback, as: as ?? key };
}

/** Lower-case a string field. */
export function lower(key: string, as?: string): LowerDef {
	return { type: "lower", key, as: as ?? key };
}

/** Apply a custom extractor function. */
export function custom(as: string, fn: (item: Record<string, unknown>) => unknown): CustomDef {
	return { type: "custom", as, fn };
}

// ---------------------------------------------------------------------------
// Extraction engine
// ---------------------------------------------------------------------------

/**
 * Why: Centralizes schema-driven field extraction so every CLI command produces
 * structurally identical output without per-command switch statements.
 * What: Walks a SchemaDef[] and builds a plain object from item using the def rules.
 * Test: Pass a mock item with known keys and a schema using every def type; assert
 * the result object has the correct keys/values for each type.
 */
export function extract(
	item: Record<string, unknown>,
	schema: SchemaDef[]
): Record<string, unknown> {
	const result: Record<string, unknown> = {};

	for (const def of schema) {
		const outputKey = def.as;

		switch (def.type) {
			case "field":
				result[outputKey] = item[def.key] ?? null;
				break;

			case "pluck": {
				const pluckSrc = item[def.key] as Record<string, unknown> | null | undefined;
				result[outputKey] = pluckSrc?.[def.subkey] ?? null;
				break;
			}

			case "joinArray": {
				const arr = item[def.key];
				if (Array.isArray(arr) && arr.length > 0) {
					result[outputKey] = arr
						.map((x: unknown) =>
							typeof x === "string" ? x : (x as Record<string, unknown>)[def.subkey]
						)
						.join(",");
				} else {
					result[outputKey] = def.empty;
				}
				break;
			}

			case "relativeTime":
				result[outputKey] = formatRelativeTime(item[def.key]);
				break;

			case "boolYesNo":
				result[outputKey] = item[def.key] ? "yes" : "no";
				break;

			case "mapEnum": {
				const val = item[def.key];
				if (typeof val === "string" && val !== "" && val in def.map) {
					result[outputKey] = def.map[val];
				} else {
					result[outputKey] = def.fallback ?? val ?? "none";
				}
				break;
			}

			case "lower": {
				const lowerSrc = item[def.key];
				result[outputKey] = typeof lowerSrc === "string" ? lowerSrc.toLowerCase() : lowerSrc;
				break;
			}

			case "custom":
				result[outputKey] = def.fn(item);
				break;

			default: {
				const _exhaustive: never = def;
				throw new Error(`Unknown schema def type: ${JSON.stringify(_exhaustive)}`);
			}
		}
	}

	return result;
}

// ---------------------------------------------------------------------------
// Render helpers
// ---------------------------------------------------------------------------

/**
 * Why: Provides a single render path for list results across all @axi-office CLIs.
 * What: Extracts each item via schema then encodes the result array as TOON.
 * Test: Pass a label and single-item array; assert the output string contains the label.
 */
export function renderList(
	label: string,
	items: Record<string, unknown>[],
	schema: SchemaDef[]
): string {
	const extracted = items.map((item) => extract(item, schema));
	return encode({ [label]: extracted });
}

/**
 * Why: Provides a single-item detail render path used by "get" commands.
 * What: Extracts a single item via schema then encodes the result as TOON.
 * Test: Pass a label and a mock item; assert the encoded output contains the label key.
 */
export function renderDetail(
	label: string,
	item: Record<string, unknown>,
	schema: SchemaDef[]
): string {
	const extracted = extract(item, schema);
	return encode({ [label]: extracted });
}

/**
 * Why: Provides a standard help suggestion block for AXI CLI home and error outputs.
 * What: Formats help lines with leading indentation into a TOON help block.
 * Test: Pass ["foo", "bar"]; assert the returned string starts with "help[2]:" and
 * contains both indented lines.
 */
export function renderHelp(lines: string[]): string {
	if (lines.length === 0) return "";
	const indented = lines.map((l) => `  ${l}`).join("\n");
	return `help[${lines.length}]:\n${indented}`;
}

/** Combine multiple TOON blocks into a single output string, filtering empty values. */
export function renderOutput(blocks: (string | undefined | null)[]): string {
	return blocks.filter(Boolean).join("\n");
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Why: Produces token-efficient relative timestamps without external date libraries.
 * What: Converts an ISO date string into a human-readable age string (e.g. "3h ago").
 * Test: Pass an ISO string 90 minutes ago; assert the result is "1h ago".
 */
export function formatRelativeTime(iso: unknown): string {
	if (!iso || typeof iso !== "string") return "unknown";
	const then = new Date(iso).getTime();
	if (Number.isNaN(then)) return "unknown";

	const diffMs = Date.now() - then;
	const diffSec = Math.floor(diffMs / 1000);

	if (diffSec < 60) return "just now";
	const diffMin = Math.floor(diffSec / 60);
	if (diffMin < 60) return `${diffMin}m ago`;
	const diffHr = Math.floor(diffMin / 60);
	if (diffHr < 24) return `${diffHr}h ago`;
	const diffDay = Math.floor(diffHr / 24);
	if (diffDay < 30) return `${diffDay}d ago`;
	const diffMon = Math.floor(diffDay / 30);
	if (diffMon < 12) return `${diffMon}mo ago`;
	const diffYr = Math.floor(diffMon / 12);
	return `${diffYr}y ago`;
}
