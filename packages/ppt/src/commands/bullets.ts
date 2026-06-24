/**
 * Why: Bullet points are the canonical way to present lists; this maps `bullets` to
 * add_bullet_points, accepting a JSON array of strings as the points.
 * What: Validates <file> <slide-index> <placeholder-idx> <points-json>, parses the
 * JSON array, and calls add_bullet_points via withOpenSave.
 * Test: Mock the client, call bulletsCommand(["/tmp/x.pptx", "0", "1", '["A","B"]']),
 * assert callTool called with "add_bullet_points" and bullet_points === ["A","B"].
 * Also: pass invalid JSON and assert AxiError VALIDATION_ERROR.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { withOpenSave } from "../session.js";

export async function bulletsCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, slideIndexRaw, placeholderIdxRaw, pointsJson] = positionals;
	if (
		!file ||
		slideIndexRaw === undefined ||
		placeholderIdxRaw === undefined ||
		pointsJson === undefined
	) {
		throw new AxiError(
			"file, slide-index, placeholder-idx and points-json are required",
			"VALIDATION_ERROR",
			[
				"ppt-axi bullets <file> <slide-index> <placeholder-idx> <points-json>",
				"",
				'  <points-json>  JSON array of strings, e.g. \'["Point 1","Point 2"]\'',
			]
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

	let bulletPoints: unknown;
	try {
		bulletPoints = JSON.parse(pointsJson);
	} catch {
		throw new AxiError("points-json is not valid JSON", "VALIDATION_ERROR", [
			'Example: \'["Point 1","Point 2"]\'',
		]);
	}
	if (!Array.isArray(bulletPoints)) {
		throw new AxiError("points-json must be a JSON array of strings", "VALIDATION_ERROR", [
			'Example: \'["Point 1","Point 2"]\'',
		]);
	}

	return withOpenSave(file, async (client) => {
		return client.callTool("add_bullet_points", {
			slide_index: slideIndex,
			placeholder_idx: placeholderIdx,
			bullet_points: bulletPoints,
		});
	});
}
