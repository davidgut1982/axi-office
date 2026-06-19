/**
 * Why: The word commands write and read real .docx files, so the tests exercise the full
 * docx + mammoth round trip rather than mocking — this is the only way to prove the
 * documents are actually well-formed and the patch substitution works.
 * What: Creates a temp dir, runs create/from-markdown/patch, reads results back via the
 * read/info commands, and asserts the expected content. Cleans up the temp dir.
 * Test: This file is the test.
 */
import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { AxiError } from "@axi-office/core";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { createCommand } from "../src/commands/create.js";
import { fromMarkdownCommand } from "../src/commands/from-markdown.js";
import { infoCommand } from "../src/commands/info.js";
import { patchCommand } from "../src/commands/patch.js";
import { readCommand } from "../src/commands/read.js";
import { parseInline, parseMarkdown } from "../src/markdown.js";

let dir: string;

beforeAll(() => {
	dir = mkdtempSync(join(tmpdir(), "word-axi-test-"));
});

afterAll(() => {
	rmSync(dir, { recursive: true, force: true });
});

describe("create + read round trip", () => {
	it("creates a docx and reads its heading back", async () => {
		const out = join(dir, "created.docx");
		const spec = JSON.stringify({
			title: "My Doc",
			sections: [
				{ type: "heading", level: 1, text: "Introduction" },
				{ type: "paragraph", text: "Hello world" },
				{ type: "list", items: ["Item 1", "Item 2"] },
				{
					type: "table",
					rows: [
						["H1", "H2"],
						["V1", "V2"],
					],
				},
			],
		});

		const result = (await createCommand([out, spec])) as { created: string };
		expect(result.created).toBe(out);
		expect(existsSync(out)).toBe(true);

		const read = (await readCommand([out])) as { text: string };
		expect(read.text).toContain("Introduction");
		expect(read.text).toContain("Hello world");
		expect(read.text).toContain("Item 1");
	});

	it("rejects a spec without sections", async () => {
		await expect(createCommand([join(dir, "bad.docx"), "{}"])).rejects.toBeInstanceOf(AxiError);
	});
});

describe("info", () => {
	it("reports counts for a created doc", async () => {
		const out = join(dir, "info.docx");
		await createCommand([
			out,
			JSON.stringify({
				sections: [
					{ type: "heading", level: 1, text: "Title" },
					{ type: "paragraph", text: "one two three four" },
				],
			}),
		]);
		const info = (await infoCommand([out])) as {
			words: number;
			paragraphs: number;
		};
		expect(info.words).toBeGreaterThan(0);
		expect(info.paragraphs).toBeGreaterThanOrEqual(1);
	});
});

describe("from-markdown", () => {
	it("converts markdown headings and lists to a readable docx", async () => {
		const md = join(dir, "in.md");
		const out = join(dir, "from-md.docx");
		writeFileSync(md, "# Heading One\n\nSome **bold** text\n\n- a\n- b\n");

		await fromMarkdownCommand([md, out]);
		const read = (await readCommand([out])) as { text: string };
		expect(read.text).toContain("Heading One");
		expect(read.text).toContain("bold");
		expect(read.text).toContain("a");
	});
});

describe("patch", () => {
	it("replaces {{key}} placeholders", async () => {
		const template = join(dir, "template.docx");
		const out = join(dir, "patched.docx");
		await createCommand([
			template,
			JSON.stringify({
				sections: [{ type: "paragraph", text: "{{name}}" }],
			}),
		]);

		const result = (await patchCommand([
			template,
			JSON.stringify({ name: "Ada Lovelace" }),
			"--out",
			out,
		])) as { keys: string[] };
		expect(result.keys).toContain("name");

		const read = (await readCommand([out])) as { text: string };
		expect(read.text).toContain("Ada Lovelace");
		expect(read.text).not.toContain("{{name}}");
	});

	it("rejects non-object data json", async () => {
		await expect(patchCommand([join(dir, "template.docx"), "[1,2]"])).rejects.toBeInstanceOf(
			AxiError
		);
	});
});

describe("markdown parser", () => {
	it("parses inline bold", () => {
		const runs = parseInline("a **b** c");
		expect(runs).toHaveLength(3);
		expect(runs[1]).toEqual({ text: "b", bold: true });
	});

	it("groups consecutive bullets into one list", () => {
		const spec = parseMarkdown("- one\n- two");
		expect(spec.sections).toHaveLength(1);
		expect(spec.sections[0]).toEqual({ type: "list", items: ["one", "two"] });
	});
});
