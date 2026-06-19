import { describe, expect, it } from "vitest";
import {
	boolYesNo,
	custom,
	extract,
	field,
	formatRelativeTime,
	joinArray,
	lower,
	mapEnum,
	pluck,
	relativeTime,
	renderDetail,
	renderHelp,
	renderList,
	renderOutput,
} from "../src/toon.js";

// ---------------------------------------------------------------------------
// Schema constructors
// ---------------------------------------------------------------------------

describe("field()", () => {
	it("uses key as output label when as is omitted", () => {
		const def = field("name");
		expect(def).toEqual({ type: "field", key: "name", as: "name" });
	});

	it("uses explicit as label", () => {
		const def = field("name", "displayName");
		expect(def).toEqual({ type: "field", key: "name", as: "displayName" });
	});
});

describe("pluck()", () => {
	it("defaults as to subkey", () => {
		const def = pluck("author", "login");
		expect(def).toEqual({
			type: "pluck",
			key: "author",
			subkey: "login",
			as: "login",
		});
	});
});

describe("boolYesNo()", () => {
	it("creates a boolYesNo def", () => {
		expect(boolYesNo("active")).toEqual({
			type: "boolYesNo",
			key: "active",
			as: "active",
		});
	});
});

describe("mapEnum()", () => {
	it("creates a mapEnum def", () => {
		const map = { OPEN: "open", CLOSED: "closed" };
		expect(mapEnum("state", map, "unknown", "status")).toEqual({
			type: "mapEnum",
			key: "state",
			map,
			fallback: "unknown",
			as: "status",
		});
	});
});

describe("custom()", () => {
	it("stores the function", () => {
		const fn = (item: Record<string, unknown>) => String(item.x);
		const def = custom("label", fn);
		expect(def.type).toBe("custom");
		expect(def.as).toBe("label");
		expect(def.fn).toBe(fn);
	});
});

// ---------------------------------------------------------------------------
// extract()
// ---------------------------------------------------------------------------

describe("extract()", () => {
	const item = {
		id: "abc",
		score: 42,
		author: { login: "alice" },
		tags: [{ name: "ts" }, { name: "node" }],
		createdAt: new Date(Date.now() - 90 * 60 * 1000).toISOString(), // 90 min ago
		active: true,
		state: "OPEN",
		title: "  HELLO WORLD  ",
	};

	it("extracts a simple field", () => {
		const result = extract(item, [field("id")]);
		expect(result.id).toBe("abc");
	});

	it("extracts a nested pluck", () => {
		const result = extract(item, [pluck("author", "login")]);
		expect(result.login).toBe("alice");
	});

	it("joins an array field", () => {
		const result = extract(item, [joinArray("tags", "name", "tags")]);
		expect(result.tags).toBe("ts,node");
	});

	it("returns empty sentinel for empty array", () => {
		const result = extract({ tags: [] }, [
			joinArray("tags", "name", "tags", "—"),
		]);
		expect(result.tags).toBe("—");
	});

	it("formats relativeTime for ~90 minutes ago", () => {
		const result = extract(item, [relativeTime("createdAt", "age")]);
		expect(result.age).toBe("1h ago");
	});

	it("renders boolYesNo true as 'yes'", () => {
		const result = extract(item, [boolYesNo("active")]);
		expect(result.active).toBe("yes");
	});

	it("renders boolYesNo false as 'no'", () => {
		const result = extract({ active: false }, [boolYesNo("active")]);
		expect(result.active).toBe("no");
	});

	it("maps enum via mapEnum", () => {
		const result = extract(item, [
			mapEnum("state", { OPEN: "open", CLOSED: "closed" }, "unknown", "status"),
		]);
		expect(result.status).toBe("open");
	});

	it("uses fallback for unmapped enum value", () => {
		const result = extract({ state: "MERGED" }, [
			mapEnum("state", { OPEN: "open" }, "other", "status"),
		]);
		expect(result.status).toBe("other");
	});

	it("lowercases a field via lower()", () => {
		const result = extract({ title: "HELLO" }, [lower("title")]);
		expect(result.title).toBe("hello");
	});

	it("applies a custom extractor", () => {
		const result = extract({ score: 42 }, [
			custom("grade", (i) => ((i.score as number) >= 40 ? "A" : "B")),
		]);
		expect(result.grade).toBe("A");
	});

	it("returns null for missing field", () => {
		const result = extract({}, [field("missing")]);
		expect(result.missing).toBeNull();
	});
});

// ---------------------------------------------------------------------------
// renderList / renderDetail
// ---------------------------------------------------------------------------

describe("renderList()", () => {
	it("returns a non-empty string containing the label", () => {
		const items = [{ id: "1", name: "foo" }];
		const schema = [field("id"), field("name")];
		const output = renderList("items", items, schema);
		expect(typeof output).toBe("string");
		expect(output.length).toBeGreaterThan(0);
		// TOON encode includes the label key
		expect(output).toContain("items");
	});
});

describe("renderDetail()", () => {
	it("returns a non-empty string containing the label", () => {
		const item = { id: "1" };
		const output = renderDetail("entry", item, [field("id")]);
		expect(output).toContain("entry");
	});
});

// ---------------------------------------------------------------------------
// renderHelp()
// ---------------------------------------------------------------------------

describe("renderHelp()", () => {
	it("returns empty string for no lines", () => {
		expect(renderHelp([])).toBe("");
	});

	it("formats a help block with count and indented lines", () => {
		const output = renderHelp(["do this", "or that"]);
		expect(output).toMatch(/^help\[2\]:/);
		expect(output).toContain("  do this");
		expect(output).toContain("  or that");
	});
});

// ---------------------------------------------------------------------------
// renderOutput()
// ---------------------------------------------------------------------------

describe("renderOutput()", () => {
	it("joins non-empty blocks", () => {
		const out = renderOutput(["a", "b", null, undefined, "c"]);
		expect(out).toBe("a\nb\nc");
	});
});

// ---------------------------------------------------------------------------
// formatRelativeTime()
// ---------------------------------------------------------------------------

describe("formatRelativeTime()", () => {
	const ago = (ms: number) => new Date(Date.now() - ms).toISOString();

	it("returns 'just now' for <60 seconds ago", () => {
		expect(formatRelativeTime(ago(10_000))).toBe("just now");
	});

	it("returns 'Xm ago' for <1 hour", () => {
		expect(formatRelativeTime(ago(5 * 60_000))).toBe("5m ago");
	});

	it("returns 'Xh ago' for <24 hours", () => {
		expect(formatRelativeTime(ago(3 * 3_600_000))).toBe("3h ago");
	});

	it("returns 'Xd ago' for <30 days", () => {
		expect(formatRelativeTime(ago(7 * 86_400_000))).toBe("7d ago");
	});

	it("returns 'unknown' for null/undefined", () => {
		expect(formatRelativeTime(null)).toBe("unknown");
		expect(formatRelativeTime(undefined)).toBe("unknown");
	});

	it("returns 'unknown' for invalid ISO string", () => {
		expect(formatRelativeTime("not-a-date")).toBe("unknown");
	});
});
