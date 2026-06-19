/**
 * Why: Template-driven document generation (mail-merge style) needs placeholder
 * substitution; `patch` replaces {{key}} markers in a .docx with provided values.
 * What: Validates <in.docx> <data-json>, builds a docx patch set keyed by placeholder, and
 * writes the result to --out or back over the input file.
 * Test: Create a docx containing "{{name}}", run patchCommand([file, '{"name":"Ada"}']),
 * read it back via mammoth, assert "Ada" appears and "{{name}}" does not.
 */
import { readFileSync, writeFileSync } from "node:fs";
import { AxiError, parseFlags } from "@axi-office/core";
import { type IPatch, PatchType, TextRun, patchDocument } from "docx";
import { resolveInBase } from "../paths.js";

export async function patchCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args);
	const [file, dataJson] = positionals;
	if (!file || !dataJson) {
		throw new AxiError(
			"input.docx and data-json are required",
			"VALIDATION_ERROR",
			["word-axi patch <in.docx> <data-json> [--out FILE]"],
		);
	}

	const baseDir =
		typeof flags["base-dir"] === "string" ? flags["base-dir"] : undefined;
	const resolvedFile = resolveInBase(baseDir, file);

	let data: unknown;
	try {
		data = JSON.parse(dataJson);
	} catch {
		throw new AxiError("data-json is not valid JSON", "VALIDATION_ERROR");
	}
	if (typeof data !== "object" || data === null || Array.isArray(data)) {
		throw new AxiError("data-json must be a JSON object", "VALIDATION_ERROR");
	}

	const patches: Record<string, IPatch> = {};
	for (const [key, value] of Object.entries(data as Record<string, unknown>)) {
		patches[key] = {
			type: PatchType.PARAGRAPH,
			children: [new TextRun(String(value))],
		};
	}

	let input: Buffer;
	try {
		input = readFileSync(resolvedFile);
	} catch {
		throw new AxiError(`cannot read docx file: ${resolvedFile}`, "IO_ERROR");
	}

	const patched = await patchDocument({
		outputType: "nodebuffer",
		data: input,
		patches,
	});

	const outRaw = typeof flags.out === "string" ? flags.out : file;
	const resolvedOut = resolveInBase(baseDir, outRaw);
	writeFileSync(resolvedOut, patched);

	return { patched: resolvedOut, keys: Object.keys(patches) };
}
