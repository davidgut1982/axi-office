/**
 * Why: Inspecting a presentation's metadata (slide count, title, etc.) before editing
 * is a common agent workflow step; this maps `info` to get_presentation_info.
 * What: Validates <file>, opens read-only, and calls get_presentation_info with the
 * backend's "current presentation" default (no presentation_id).
 * Test: Mock the client, call infoCommand(["/tmp/x.pptx"]), assert callTool was
 * invoked with "open_presentation" and then "get_presentation_info" (no save).
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { withOpenReadonly } from "../session.js";

export async function infoCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const file = positionals[0];
	if (!file) {
		throw new AxiError("file path is required", "VALIDATION_ERROR", ["ppt-axi info <file>"]);
	}

	return withOpenReadonly(file, async (client) => {
		return client.callTool("get_presentation_info", {});
	});
}
