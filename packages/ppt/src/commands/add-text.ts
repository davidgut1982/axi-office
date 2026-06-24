/**
 * Why: Adding free-form text boxes at precise positions is essential for custom layouts;
 * this maps `add-text` to manage_text with operation="add".
 * What: Validates <file> <slide-index> <text>, parses optional position/size/font flags,
 * parses --color "R,G,B" into [int,int,int], and calls manage_text via withOpenSave
 * with presentation_id.
 * Test: Mock the client, call addTextCommand(["/tmp/x.pptx","0","Hello","--left","1",
 * "--top","2"]), assert callTool called with manage_text and operation="add",
 * slide_index=0, left=1.0, top=2.0. Also: pass --color "255,0,0" and assert
 * color is [255,0,0].
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { call, withOpenSave } from "../session.js";
import { parseColor } from "../utils.js";

export async function addTextCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args, ["bold", "italic", "underline"]);
	const [file, slideIndexRaw, text] = positionals;
	if (!file || slideIndexRaw === undefined || text === undefined) {
		throw new AxiError("file, slide-index and text are required", "VALIDATION_ERROR", [
			"ppt-axi add-text <file> <slide-index> <text> [--left=1] [--top=1] [--width=4] [--height=2]",
			"  [--font-size N] [--font-name N] [--bold] [--italic] [--underline]",
			"  [--color R,G,B] [--align left|center|right]",
		]);
	}

	const slideIndex = Number.parseInt(slideIndexRaw, 10);
	if (!Number.isFinite(slideIndex) || slideIndex < 0) {
		throw new AxiError("slide-index must be a non-negative integer", "VALIDATION_ERROR");
	}

	const left = Number.parseFloat(typeof flags.left === "string" ? flags.left : "1");
	const top = Number.parseFloat(typeof flags.top === "string" ? flags.top : "1");
	const width = Number.parseFloat(typeof flags.width === "string" ? flags.width : "4");
	const height = Number.parseFloat(typeof flags.height === "string" ? flags.height : "2");

	return withOpenSave(file, async (client, presentationId) => {
		const toolArgs: Record<string, unknown> = {
			operation: "add",
			slide_index: slideIndex,
			text,
			left,
			top,
			width,
			height,
			presentation_id: presentationId,
		};
		if (typeof flags["font-size"] === "string") {
			const fontSize = Number.parseInt(flags["font-size"], 10);
			if (!Number.isFinite(fontSize) || fontSize <= 0) {
				throw new AxiError("--font-size must be a positive integer", "VALIDATION_ERROR", [
					"Example: --font-size 24",
				]);
			}
			toolArgs.font_size = fontSize;
		}
		if (typeof flags["font-name"] === "string") toolArgs.font_name = flags["font-name"];
		if (flags.bold === true) toolArgs.bold = true;
		if (flags.italic === true) toolArgs.italic = true;
		if (flags.underline === true) toolArgs.underline = true;
		if (typeof flags.color === "string") toolArgs.color = parseColor(flags.color, "--color");
		if (typeof flags.align === "string") toolArgs.align = flags.align;

		return call(client, "manage_text", toolArgs);
	});
}
