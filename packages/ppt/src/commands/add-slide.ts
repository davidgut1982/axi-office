/**
 * Why: Slides are the fundamental content unit; this maps `add-slide` to the add_slide
 * MCP tool which appends a slide with the given layout, optional title, and color scheme.
 * What: Validates <file>, parses optional --layout, --title, --color-scheme flags,
 * and calls add_slide via withOpenSave with only provided optional fields, plus presentation_id.
 * Test: Mock the client, call addSlideCommand(["/tmp/x.pptx"]), assert callTool was
 * invoked with add_slide and layout_index=1 (default). Also: pass --title "Hello" and
 * assert title is included in the tool args.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { call, withOpenSave } from "../session.js";

export async function addSlideCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args);
	const file = positionals[0];
	if (!file) {
		throw new AxiError("file path is required", "VALIDATION_ERROR", [
			"ppt-axi add-slide <file> [--layout N] [--title T] [--color-scheme S]",
		]);
	}

	const layoutRaw = typeof flags.layout === "string" ? flags.layout : "1";
	const layoutIndex = Number.parseInt(layoutRaw, 10);
	if (!Number.isFinite(layoutIndex) || layoutIndex < 0) {
		throw new AxiError("--layout must be a non-negative integer", "VALIDATION_ERROR", [
			"Example: --layout 1",
		]);
	}

	return withOpenSave(file, async (client, presentationId) => {
		const toolArgs: Record<string, unknown> = {
			layout_index: layoutIndex,
			presentation_id: presentationId,
		};
		if (typeof flags.title === "string") toolArgs.title = flags.title;
		if (typeof flags["color-scheme"] === "string") toolArgs.color_scheme = flags["color-scheme"];
		return call(client, "add_slide", toolArgs);
	});
}
