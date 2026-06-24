/**
 * Why: Auto-generating a complete presentation from a topic is the highest-value
 * command for agents; this maps `auto-generate` to auto_generate_presentation.
 * What: Validates <file> and <topic>, parses optional flags for slide count, type,
 * color scheme, and media options, validates enum values, then calls
 * auto_generate_presentation via withCreateSave with presentation_id.
 * Test: Mock the client, call autoGenerateCommand(["/tmp/x.pptx","AI trends"]), assert
 * callTool called with "create_presentation" then "auto_generate_presentation" then
 * "save_presentation". Also: pass invalid --type and assert AxiError VALIDATION_ERROR.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { call, withCreateSave } from "../session.js";

const VALID_TYPES = ["business", "academic", "creative"] as const;
const VALID_COLOR_SCHEMES = ["modern_blue", "corporate_gray", "elegant_green", "warm_red"] as const;

export async function autoGenerateCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args, ["no-charts", "images"]);
	const [file, topic] = positionals;
	if (!file || topic === undefined) {
		throw new AxiError("file and topic are required", "VALIDATION_ERROR", [
			"ppt-axi auto-generate <file> <topic> [--slides=5] [--type=business] [--color-scheme=modern_blue]",
			"  [--no-charts] [--images]",
			`  --type:          one of ${VALID_TYPES.join(", ")}`,
			`  --color-scheme:  one of ${VALID_COLOR_SCHEMES.join(", ")}`,
		]);
	}

	const presentationType = typeof flags.type === "string" ? flags.type : "business";
	if (!VALID_TYPES.includes(presentationType as (typeof VALID_TYPES)[number])) {
		throw new AxiError(`--type "${presentationType}" is not supported`, "VALIDATION_ERROR", [
			`Supported types: ${VALID_TYPES.join(", ")}`,
		]);
	}

	const colorScheme =
		typeof flags["color-scheme"] === "string" ? flags["color-scheme"] : "modern_blue";
	if (!VALID_COLOR_SCHEMES.includes(colorScheme as (typeof VALID_COLOR_SCHEMES)[number])) {
		throw new AxiError(`--color-scheme "${colorScheme}" is not supported`, "VALIDATION_ERROR", [
			`Supported schemes: ${VALID_COLOR_SCHEMES.join(", ")}`,
		]);
	}

	const slidesRaw = typeof flags.slides === "string" ? flags.slides : "5";
	const slideCount = Number.parseInt(slidesRaw, 10);
	if (!Number.isFinite(slideCount) || slideCount < 1) {
		throw new AxiError("--slides must be a positive integer", "VALIDATION_ERROR");
	}

	// include_charts defaults to true unless --no-charts is explicitly set
	const includeCharts = flags["no-charts"] !== true;
	const includeImages = flags.images === true ? true : undefined;

	return withCreateSave(file, async (client, presentationId) => {
		const toolArgs: Record<string, unknown> = {
			topic,
			slide_count: slideCount,
			presentation_type: presentationType,
			color_scheme: colorScheme,
			include_charts: includeCharts,
			presentation_id: presentationId,
		};
		if (includeImages !== undefined) toolArgs.include_images = includeImages;

		return call(client, "auto_generate_presentation", toolArgs);
	});
}
