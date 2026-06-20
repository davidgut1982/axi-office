/**
 * Why: Applying styles to a range improves spreadsheet readability; this maps
 * `format-range` to the format_range MCP tool (haris-musa backend).
 * What: Validates <file> <sheet> <start-cell> <format-json>, parses a flat style
 * object (bold, italic, font_size, font_color, bg_color, etc.), and forwards it to
 * format_range. Optional end-cell widens the range.
 * Note: The haris-musa tool takes flat style properties directly, NOT the negokaz
 * 2D per-cell array. <format-json> must be a plain object, not a 2D array.
 * Test: Mock the client, call formatRangeCommand(["/tmp/x.xlsx", "Sheet1", "A1",
 * '{"bold":true}', "B1"]), assert callTool was invoked with "format_range" and
 * the spread style keys alongside filepath/sheet_name/start_cell/end_cell.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function formatRangeCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, startCell, formatJson, endCell] = positionals;
	if (!file || !sheet || !startCell || !formatJson) {
		throw new AxiError(
			"file, sheet, start-cell and format-json are required",
			"VALIDATION_ERROR",
			[
				"excel-axi format-range <file> <sheet> <start-cell> <format-json> [end-cell]",
				"",
				"<format-json> is a flat JSON object with style properties:",
				"  bold (bool), italic (bool), underline (bool)",
				"  font_size (int), font_color (str), bg_color (str)",
				"  border_style (str), border_color (str)",
				"  number_format (str), alignment (str), wrap_text (bool)",
				"",
				'Example: \'{"bold":true,"font_size":14,"bg_color":"FFFF00"}\'',
			]
		);
	}

	let styleObj: unknown;
	try {
		styleObj = JSON.parse(formatJson);
	} catch {
		throw new AxiError("format-json is not valid JSON", "VALIDATION_ERROR", [
			'Example: \'{"bold":true,"font_color":"FF0000"}\'',
		]);
	}
	if (typeof styleObj !== "object" || styleObj === null || Array.isArray(styleObj)) {
		throw new AxiError(
			"format-json must be a plain object with style properties (not an array)",
			"VALIDATION_ERROR",
			[
				"Supported keys: bold, italic, underline, font_size, font_color, bg_color,",
				"  border_style, border_color, number_format, alignment, wrap_text",
				'Example: \'{"bold":true,"font_size":14}\'',
			]
		);
	}

	const toolArgs: Record<string, unknown> = {
		filepath: file,
		sheet_name: sheet,
		start_cell: startCell,
		...(styleObj as Record<string, unknown>),
	};
	if (endCell) toolArgs.end_cell = endCell;

	return getClient().callTool("format_range", toolArgs);
}
