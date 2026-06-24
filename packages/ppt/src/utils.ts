/**
 * Why: parseColor is used in add-text and add-shape with the same logic but previously
 * existed as duplicate local functions with diverged signatures. Centralizing it
 * prevents drift and makes the constraint (0-255 per channel) a single source of truth.
 * What: Parses a "R,G,B" string into a typed [number, number, number] tuple, throwing
 * AxiError VALIDATION_ERROR with a hint showing the flag name on any parse failure.
 * Test: parseColor("255,0,128", "--color") → [255, 0, 128];
 * parseColor("notacolor", "--color") → throws AxiError VALIDATION_ERROR.
 */
import { AxiError } from "@axi-office/core";

export function parseColor(colorStr: string, flagName: string): [number, number, number] {
	const parts = colorStr.split(",").map((s) => Number.parseInt(s.trim(), 10));
	if (parts.length !== 3 || parts.some((n) => !Number.isFinite(n) || n < 0 || n > 255)) {
		throw new AxiError(
			`${flagName} "${colorStr}" is not a valid R,G,B triplet`,
			"VALIDATION_ERROR",
			[`Example: ${flagName} 255,0,0`]
		);
	}
	return parts as [number, number, number];
}

/**
 * Why: Numeric position/size flags (--left/--top/--width/--height) feed straight into the
 * backend; a non-numeric value yields NaN, which JSON-serializes to null over the MCP
 * transport and produces a confusing backend error instead of a clean VALIDATION_ERROR.
 * What: Parses a float flag, returning defaultValue when the flag is absent and throwing
 * AxiError VALIDATION_ERROR when present but non-finite. With no defaultValue, returns
 * undefined for absent flags (for flags that are only forwarded when supplied).
 * Test: parseFloatFlag("2.5", "--left", 1) → 2.5; parseFloatFlag(undefined, "--left", 1) → 1;
 * parseFloatFlag("foo", "--left", 1) → throws AxiError VALIDATION_ERROR;
 * parseFloatFlag(undefined, "--width") → undefined.
 */
export function parseFloatFlag(value: unknown, flagName: string, defaultValue: number): number;
export function parseFloatFlag(
	value: unknown,
	flagName: string,
	defaultValue?: undefined
): number | undefined;
export function parseFloatFlag(
	value: unknown,
	flagName: string,
	defaultValue?: number
): number | undefined {
	if (typeof value !== "string") return defaultValue;
	const parsed = Number.parseFloat(value);
	if (!Number.isFinite(parsed)) {
		throw new AxiError(`${flagName} "${value}" is not a valid number`, "VALIDATION_ERROR", [
			`Example: ${flagName} 1.5`,
		]);
	}
	return parsed;
}
