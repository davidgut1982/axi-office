/**
 * Why: Provides a minimal, dependency-free flag parser for AXI command handlers so
 * every @axi-office CLI can parse --key value, --key=value, and boolean flags without
 * pulling in a heavy argument parsing library.
 * What: Exports parseFlags, parseLimit, parseKeyValuePairs, coerceValue, and string
 * utility helpers ported from lore-axi's args.ts.
 * Test: Call parseFlags(["--limit", "10", "search", "--verbose"]) and assert
 * positionals=["search"], flags.limit="10", flags.verbose=true.
 */

/**
 * Result of parseFlags — positionals are non-flag tokens, flags are --key value pairs.
 */
export interface ParsedArgs {
  positionals: string[];
  flags: Record<string, string | true>;
}

/**
 * Why: Centralizes flag parsing so command handlers stay simple.
 * What: Parses --key value, --key=value, and boolean --flag forms. Names listed in
 * booleans are always treated as boolean flags even when a non-flag token follows.
 * Test: parseFlags(["--out", "file.txt", "--dry-run"], ["dry-run"]) →
 * { positionals: [], flags: { out: "file.txt", "dry-run": true } }.
 */
export function parseFlags(args: string[], booleans: string[] = []): ParsedArgs {
  const positionals: string[] = [];
  const flags: Record<string, string | true> = {};

  for (let i = 0; i < args.length; i++) {
    const token = args[i];
    if (!token.startsWith("--")) {
      positionals.push(token);
      continue;
    }

    const body = token.slice(2);
    const eq = body.indexOf("=");
    if (eq !== -1) {
      flags[body.slice(0, eq)] = body.slice(eq + 1);
      continue;
    }

    const next = args[i + 1];
    if (next !== undefined && !next.startsWith("--") && !booleans.includes(body)) {
      flags[body] = next;
      i++;
    } else {
      flags[body] = true;
    }
  }

  return { positionals, flags };
}

/**
 * Why: Limits results to a safe maximum to prevent runaway MCP tool calls.
 * What: Parses a --limit flag value into a clamped positive integer, returning
 * fallback when the value is absent or invalid.
 * Test: parseLimit("5", 10, 100) → 5; parseLimit("999", 10, 100) → 100;
 * parseLimit(undefined, 10, 100) → 10.
 */
export function parseLimit(value: unknown, fallback: number, max: number): number {
  if (typeof value !== "string") return fallback;
  const parsed = Number.parseInt(value, 10);
  if (!Number.isFinite(parsed) || parsed < 1) return fallback;
  return Math.min(parsed, max);
}

/**
 * Why: Converts positional key=value tokens into a record for MCP tool arguments.
 * What: Splits each arg on the first "=" and stores left=key, right=value.
 * Test: parseKeyValuePairs(["name=foo", "type=bar"]) → { name: "foo", type: "bar" }.
 */
export function parseKeyValuePairs(args: string[]): Record<string, string> {
  const result: Record<string, string> = {};
  for (const arg of args) {
    const eq = arg.indexOf("=");
    if (eq !== -1) {
      result[arg.slice(0, eq)] = arg.slice(eq + 1);
    }
  }
  return result;
}

/**
 * Why: Avoids stringly-typed MCP arguments by coercing common scalar strings to
 * their native types.
 * What: Converts "true"/"false" to boolean, numeric strings to number; leaves other
 * strings unchanged.
 * Test: coerceValue("true") → true; coerceValue("42") → 42; coerceValue("hi") → "hi".
 */
export function coerceValue(value: string): string | number | boolean {
  if (value === "true") return true;
  if (value === "false") return false;
  if (value === "null") return "null";
  const num = Number(value);
  if (String(num) === value && Number.isFinite(num)) return num;
  return value;
}

/**
 * Why: Keeps log and display strings clean when user input may have runs of whitespace.
 * What: Replaces all whitespace runs with a single space and trims the result.
 * Test: collapseWhitespace("  foo   bar  ") → "foo bar".
 */
export function collapseWhitespace(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

/**
 * Why: Prevents very long strings from breaking TOON table layout.
 * What: Collapses whitespace then hard-truncates to max characters, appending " …".
 * Test: truncateLine("a very long string", 10) → "a very lon …".
 */
export function truncateLine(text: string, max: number): string {
  const line = collapseWhitespace(text);
  if (line.length <= max) return line;
  return `${line.slice(0, max).trimEnd()} …`;
}
