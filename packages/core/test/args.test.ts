import { describe, expect, it } from "vitest";
import {
  coerceValue,
  collapseWhitespace,
  parseFlags,
  parseKeyValuePairs,
  parseLimit,
  truncateLine,
} from "../src/args.js";

describe("parseFlags()", () => {
  it("parses positional args", () => {
    const result = parseFlags(["foo", "bar"]);
    expect(result.positionals).toEqual(["foo", "bar"]);
    expect(result.flags).toEqual({});
  });

  it("parses --key value", () => {
    const result = parseFlags(["--limit", "10"]);
    expect(result.flags.limit).toBe("10");
    expect(result.positionals).toEqual([]);
  });

  it("parses --key=value", () => {
    const result = parseFlags(["--out=file.txt"]);
    expect(result.flags.out).toBe("file.txt");
  });

  it("parses boolean flag without value", () => {
    const result = parseFlags(["--verbose"]);
    expect(result.flags.verbose).toBe(true);
  });

  it("treats a named flag as boolean when listed in booleans", () => {
    const result = parseFlags(["--dry-run", "next-arg"], ["dry-run"]);
    expect(result.flags["dry-run"]).toBe(true);
    expect(result.positionals).toEqual(["next-arg"]);
  });

  it("handles mixed positionals and flags", () => {
    const result = parseFlags(["search", "--limit", "5", "--verbose"]);
    expect(result.positionals).toEqual(["search"]);
    expect(result.flags.limit).toBe("5");
    expect(result.flags.verbose).toBe(true);
  });

  it("handles flag followed by another flag (treats first as boolean)", () => {
    const result = parseFlags(["--foo", "--bar"]);
    expect(result.flags.foo).toBe(true);
    expect(result.flags.bar).toBe(true);
  });
});

describe("parseLimit()", () => {
  it("returns fallback for undefined", () => {
    expect(parseLimit(undefined, 10, 100)).toBe(10);
  });

  it("parses a valid limit string", () => {
    expect(parseLimit("5", 10, 100)).toBe(5);
  });

  it("clamps to max", () => {
    expect(parseLimit("999", 10, 100)).toBe(100);
  });

  it("returns fallback for non-numeric string", () => {
    expect(parseLimit("abc", 10, 100)).toBe(10);
  });

  it("returns fallback for 0 or negative", () => {
    expect(parseLimit("0", 10, 100)).toBe(10);
    expect(parseLimit("-1", 10, 100)).toBe(10);
  });
});

describe("parseKeyValuePairs()", () => {
  it("parses key=value tokens", () => {
    const result = parseKeyValuePairs(["name=foo", "type=bar"]);
    expect(result).toEqual({ name: "foo", type: "bar" });
  });

  it("ignores tokens without '='", () => {
    const result = parseKeyValuePairs(["noeq", "key=val"]);
    expect(result).toEqual({ key: "val" });
  });

  it("returns empty object for empty array", () => {
    expect(parseKeyValuePairs([])).toEqual({});
  });
});

describe("coerceValue()", () => {
  it("coerces 'true' to boolean true", () => {
    expect(coerceValue("true")).toBe(true);
  });

  it("coerces 'false' to boolean false", () => {
    expect(coerceValue("false")).toBe(false);
  });

  it("coerces numeric string to number", () => {
    expect(coerceValue("42")).toBe(42);
    expect(coerceValue("3.14")).toBe(3.14);
  });

  it("keeps plain strings as strings", () => {
    expect(coerceValue("hello")).toBe("hello");
  });

  it("keeps 'null' as the string 'null'", () => {
    expect(coerceValue("null")).toBe("null");
  });
});

describe("collapseWhitespace()", () => {
  it("collapses multiple spaces", () => {
    expect(collapseWhitespace("  foo   bar  ")).toBe("foo bar");
  });

  it("collapses newlines and tabs", () => {
    expect(collapseWhitespace("foo\t\nbar")).toBe("foo bar");
  });

  it("returns empty string for whitespace-only input", () => {
    expect(collapseWhitespace("   ")).toBe("");
  });
});

describe("truncateLine()", () => {
  it("returns string unchanged when within max", () => {
    expect(truncateLine("short", 10)).toBe("short");
  });

  it("truncates and appends ellipsis when over max", () => {
    const result = truncateLine("a very long string here", 10);
    expect(result.endsWith(" …")).toBe(true);
    // The part before " …" should be at most 10 chars (trimmed)
    expect(result.slice(0, -2).trim().length).toBeLessThanOrEqual(10);
  });

  it("collapses whitespace before measuring", () => {
    const result = truncateLine("  hello  ", 10);
    expect(result).toBe("hello");
  });
});
