/**
 * Why: Inspecting a specific slide's shapes and content before editing is essential for
 * targeted modifications; this maps `slide-info` to get_slide_info.
 * What: Validates <file> and <slide-index> (0-based int), opens read-only, and calls
 * get_slide_info with the slide_index and presentation_id.
 * Test: Mock the client, call slideInfoCommand(["/tmp/x.pptx", "0"]), assert callTool
 * called with "get_slide_info" and { slide_index: 0, presentation_id: ... }. No save_presentation call.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { call, withOpenReadonly } from "../session.js";

export async function slideInfoCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, slideIndexRaw] = positionals;
	if (!file || slideIndexRaw === undefined) {
		throw new AxiError("file and slide-index are required", "VALIDATION_ERROR", [
			"ppt-axi slide-info <file> <slide-index>",
			"  <slide-index>  0-based slide index",
		]);
	}

	const slideIndex = Number.parseInt(slideIndexRaw, 10);
	if (!Number.isFinite(slideIndex) || slideIndex < 0) {
		throw new AxiError("slide-index must be a non-negative integer", "VALIDATION_ERROR");
	}

	return withOpenReadonly(file, async (client, presentationId) => {
		return call(client, "get_slide_info", {
			slide_index: slideIndex,
			presentation_id: presentationId,
		});
	});
}
