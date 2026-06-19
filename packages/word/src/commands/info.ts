import { AxiError, parseFlags } from "@axi-office/core";
/**
 * Why: Agents often need quick document metrics (length, structure) before deciding how to
 * process a .docx; `info` returns best-effort counts.
 * What: Validates <in.docx>, extracts raw text via mammoth, and counts words, non-empty
 * paragraphs, and heading-like lines.
 * Test: Create a docx with one heading and one paragraph; run infoCommand([file]); assert
 * words > 0 and paragraphs >= 1.
 */
import mammoth from "mammoth";

export async function infoCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const file = positionals[0];
	if (!file) {
		throw new AxiError("input.docx is required", "VALIDATION_ERROR", ["word-axi info <in.docx>"]);
	}

	const { value: text } = await mammoth.extractRawText({ path: file });
	const lines = text.split(/\r?\n/).map((l) => l.trim());
	const nonEmpty = lines.filter((l) => l.length > 0);
	const words = text.split(/\s+/).filter((w) => w.length > 0).length;
	// Heading heuristic: short, title-cased-ish lines with no terminal punctuation.
	const headings = nonEmpty.filter((l) => l.length <= 80 && !/[.!?,;:]$/.test(l)).length;

	return {
		file,
		words,
		paragraphs: nonEmpty.length,
		headings,
	};
}
