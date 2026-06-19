/**
 * Why: Agents need to generate a .docx from structured data; `create` accepts a JSON spec
 * (or stdin) and writes the document to disk.
 * What: Validates <out> and <spec-json|->, parses the spec, builds the docx, writes it.
 * Test: Call createCommand(["/tmp/o.docx", '{"sections":[{"type":"paragraph","text":"hi"}]}']),
 * assert the file exists and is non-empty.
 */
import { readFileSync, writeFileSync } from "node:fs";
import { AxiError, parseFlags } from "@axi-office/core";
import { type DocSpec, buildDocxBuffer } from "../docx-build.js";
import { resolveInBase } from "../paths.js";

/**
 * Why: Specs may be passed inline or piped; "-" means read stdin so large specs avoid
 * argv limits.
 * What: Returns the raw spec string from the argument or from stdin when arg === "-".
 * Test: readSpec("{}") returns "{}"; readSpec("-") reads /dev/stdin.
 */
function readSpec(arg: string): string {
	if (arg === "-") {
		return readFileSync(0, "utf8");
	}
	return arg;
}

export async function createCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args);
	const [out, specArg] = positionals;
	if (!out || !specArg) {
		throw new AxiError(
			"out path and spec-json are required",
			"VALIDATION_ERROR",
			["word-axi create <out.docx> <spec-json|->"],
		);
	}

	const baseDir =
		typeof flags["base-dir"] === "string" ? flags["base-dir"] : undefined;

	let spec: DocSpec;
	try {
		spec = JSON.parse(readSpec(specArg)) as DocSpec;
	} catch {
		throw new AxiError("spec-json is not valid JSON", "VALIDATION_ERROR");
	}
	if (!spec || !Array.isArray(spec.sections)) {
		throw new AxiError("spec must have a sections array", "VALIDATION_ERROR");
	}

	const resolvedOut = resolveInBase(baseDir, out);
	const buffer = await buildDocxBuffer(spec);
	writeFileSync(resolvedOut, buffer);

	return {
		created: resolvedOut,
		bytes: buffer.length,
		sections: spec.sections.length,
	};
}
