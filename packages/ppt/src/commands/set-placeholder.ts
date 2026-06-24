/**
 * Why: Template-based layouts use named/indexed placeholders for titles and body text;
 * this maps `set-placeholder` to populate_placeholder.
 * What: Validates <file> <slide-index> <placeholder-idx> <text>, opens with save, and
 * calls populate_placeholder with the parsed integer indices and presentation_id.
 * Test: Mock the client, call setPlaceholderCommand(["/tmp/x.pptx", "0", "0", "Hello"]),
 * assert callTool called with "populate_placeholder" and matching args. Save also called.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { call, withOpenSave } from "../session.js";

export async function setPlaceholderCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, slideIndexRaw, placeholderIdxRaw, text] = positionals;
	if (
		!file ||
		slideIndexRaw === undefined ||
		placeholderIdxRaw === undefined ||
		text === undefined
	) {
		throw new AxiError(
			"file, slide-index, placeholder-idx and text are required",
			"VALIDATION_ERROR",
			["ppt-axi set-placeholder <file> <slide-index> <placeholder-idx> <text>"]
		);
	}

	const slideIndex = Number.parseInt(slideIndexRaw, 10);
	if (!Number.isFinite(slideIndex) || slideIndex < 0) {
		throw new AxiError("slide-index must be a non-negative integer", "VALIDATION_ERROR");
	}

	const placeholderIdx = Number.parseInt(placeholderIdxRaw, 10);
	if (!Number.isFinite(placeholderIdx) || placeholderIdx < 0) {
		throw new AxiError("placeholder-idx must be a non-negative integer", "VALIDATION_ERROR");
	}

	return withOpenSave(file, async (client, presentationId) => {
		return call(client, "populate_placeholder", {
			slide_index: slideIndex,
			placeholder_idx: placeholderIdx,
			text,
			presentation_id: presentationId,
		});
	});
}
