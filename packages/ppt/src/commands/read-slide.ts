/**
 * Why: Reading a single slide's text lets agents work with targeted content without
 * retrieving the whole deck; this maps `read-slide` to extract_slide_text.
 * What: Validates <file> and <slide-index>, opens read-only, calls extract_slide_text.
 * Test: Mock the client, call readSlideCommand(["/tmp/x.pptx", "2"]), assert callTool
 * called with "extract_slide_text" and { slide_index: 2 }. No save call.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { withOpenReadonly } from "../session.js";

export async function readSlideCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, slideIndexRaw] = positionals;
	if (!file || slideIndexRaw === undefined) {
		throw new AxiError("file and slide-index are required", "VALIDATION_ERROR", [
			"ppt-axi read-slide <file> <slide-index>",
			"  <slide-index>  0-based slide index",
		]);
	}

	const slideIndex = Number.parseInt(slideIndexRaw, 10);
	if (!Number.isFinite(slideIndex) || slideIndex < 0) {
		throw new AxiError("slide-index must be a non-negative integer", "VALIDATION_ERROR");
	}

	return withOpenReadonly(file, async (client) => {
		return client.callTool("extract_slide_text", { slide_index: slideIndex });
	});
}
