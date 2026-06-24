/**
 * Why: Embedding images makes presentations visually rich; this maps `add-image` to
 * manage_image with operation="add" and source_type="file".
 * What: Validates <file> <slide-index> <image-path>, applies position defaults
 * (left=1, top=1), and conditionally includes width/height only when provided.
 * Passes presentation_id to manage_image.
 * Test: Mock the client, call addImageCommand(["/tmp/x.pptx","0","/img.png"]), assert
 * callTool called with manage_image, operation="add", image_source="/img.png",
 * source_type="file". Also: pass --width 4 and assert width=4 is in args.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { call, withOpenSave } from "../session.js";
import { parseFloatFlag } from "../utils.js";

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

	const left = parseFloatFlag(flags.left, "--left", 1);
	const top = parseFloatFlag(flags.top, "--top", 1);
	const width = parseFloatFlag(flags.width, "--width");
	const height = parseFloatFlag(flags.height, "--height");

	return withOpenSave(file, async (client, presentationId) => {
		const toolArgs: Record<string, unknown> = {
			operation: "add",
			slide_index: slideIndex,
			image_source: imagePath,
			source_type: "file",
			left,
			top,
			presentation_id: presentationId,
		};
		if (width !== undefined) toolArgs.width = width;
		if (height !== undefined) toolArgs.height = height;

		return call(client, "manage_image", toolArgs);
	});
}
