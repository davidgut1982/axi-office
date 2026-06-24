/**
 * Why: Extracting all text from a presentation (with slide numbers) is the primary way
 * agents read deck content; this maps `read` to extract_presentation_text.
 * What: Validates <file>, opens read-only, and calls extract_presentation_text with
 * include_slide_info=true for context-rich output.
 * Test: Mock the client, call readCommand(["/tmp/x.pptx"]), assert callTool was invoked
 * with "extract_presentation_text" and { include_slide_info: true }. No save call.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { withOpenReadonly } from "../session.js";

export async function readCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const file = positionals[0];
	if (!file) {
		throw new AxiError("file path is required", "VALIDATION_ERROR", ["ppt-axi read <file>"]);
	}

	return withOpenReadonly(file, async (client) => {
		return client.callTool("extract_presentation_text", { include_slide_info: true });
	});
}
