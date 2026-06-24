/**
 * Why: Embedding images makes presentations visually rich; this maps `add-image` to
 * manage_image with operation="add" and source_type="file".
 * What: Validates <file> <slide-index> <image-path>, applies position defaults
 * (left=1, top=1), and conditionally includes width/height only when provided.
 * Test: Mock the client, call addImageCommand(["/tmp/x.pptx","0","/img.png"]), assert
 * callTool called with manage_image, operation="add", image_source="/img.png",
 * source_type="file". Also: pass --width 4 and assert width=4 is in args.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { withOpenSave } from "../session.js";

export async function addImageCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args);
	const [file, slideIndexRaw, imagePath] = positionals;
	if (!file || slideIndexRaw === undefined || imagePath === undefined) {
		throw new AxiError("file, slide-index and image-path are required", "VALIDATION_ERROR", [
			"ppt-axi add-image <file> <slide-index> <image-path> [--left=1] [--top=1] [--width W] [--height H]",
		]);
	}

	const slideIndex = Number.parseInt(slideIndexRaw, 10);
	if (!Number.isFinite(slideIndex) || slideIndex < 0) {
		throw new AxiError("slide-index must be a non-negative integer", "VALIDATION_ERROR");
	}

	const left = Number.parseFloat(typeof flags.left === "string" ? flags.left : "1");
	const top = Number.parseFloat(typeof flags.top === "string" ? flags.top : "1");

	return withOpenSave(file, async (client) => {
		const toolArgs: Record<string, unknown> = {
			operation: "add",
			slide_index: slideIndex,
			image_source: imagePath,
			source_type: "file",
			left,
			top,
		};
		if (typeof flags.width === "string") toolArgs.width = Number.parseFloat(flags.width);
		if (typeof flags.height === "string") toolArgs.height = Number.parseFloat(flags.height);

		return client.callTool("manage_image", toolArgs);
	});
}
