/**
 * Why: Shapes (rectangles, ovals, arrows) are common decorative and structural slide
 * elements; this maps `add-shape` to add_shape.
 * What: Validates <file> <slide-index> <shape-type>, applies position/size defaults,
 * parses optional --text, --fill-color, --line-color flags.
 * Test: Mock the client, call addShapeCommand(["/tmp/x.pptx","0","rectangle"]),
 * assert callTool called with "add_shape" and shape_type="rectangle" with defaults.
 * Also: pass --fill-color "255,0,0" and assert fill_color is [255,0,0].
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { withOpenSave } from "../session.js";

function parseColor(colorStr: string, flagName: string): [number, number, number] {
	const parts = colorStr.split(",").map((s) => Number.parseInt(s.trim(), 10));
	if (parts.length !== 3 || parts.some((n) => !Number.isFinite(n) || n < 0 || n > 255)) {
		throw new AxiError(
			`${flagName} "${colorStr}" is not a valid R,G,B triplet`,
			"VALIDATION_ERROR",
			[`Example: ${flagName} 255,0,0`]
		);
	}
	return parts as [number, number, number];
}

export async function addShapeCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args);
	const [file, slideIndexRaw, shapeType] = positionals;
	if (!file || slideIndexRaw === undefined || shapeType === undefined) {
		throw new AxiError("file, slide-index and shape-type are required", "VALIDATION_ERROR", [
			"ppt-axi add-shape <file> <slide-index> <shape-type> [--left=1] [--top=1] [--width=2] [--height=2]",
			"  [--text T] [--fill-color R,G,B] [--line-color R,G,B]",
		]);
	}

	const slideIndex = Number.parseInt(slideIndexRaw, 10);
	if (!Number.isFinite(slideIndex) || slideIndex < 0) {
		throw new AxiError("slide-index must be a non-negative integer", "VALIDATION_ERROR");
	}

	const left = Number.parseFloat(typeof flags.left === "string" ? flags.left : "1");
	const top = Number.parseFloat(typeof flags.top === "string" ? flags.top : "1");
	const width = Number.parseFloat(typeof flags.width === "string" ? flags.width : "2");
	const height = Number.parseFloat(typeof flags.height === "string" ? flags.height : "2");

	return withOpenSave(file, async (client) => {
		const toolArgs: Record<string, unknown> = {
			slide_index: slideIndex,
			shape_type: shapeType,
			left,
			top,
			width,
			height,
		};
		if (typeof flags.text === "string") toolArgs.text = flags.text;
		if (typeof flags["fill-color"] === "string") {
			toolArgs.fill_color = parseColor(flags["fill-color"], "--fill-color");
		}
		if (typeof flags["line-color"] === "string") {
			toolArgs.line_color = parseColor(flags["line-color"], "--line-color");
		}

		return client.callTool("add_shape", toolArgs);
	});
}
