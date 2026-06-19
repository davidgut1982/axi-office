/**
 * Why: Applying styles to a range improves spreadsheet readability; this maps
 * `format-range` to the excel_format_range MCP tool.
 * What: Validates <file> <sheet> <range> <format-json>, parses the style JSON object,
 * and forwards it to excel_format_range.
 * Test: Mock the client, call formatRangeCommand(["/tmp/x.xlsx", "Sheet1", "A1:B2", '{"bold":true}']),
 * assert callTool was invoked with "excel_format_range" and a parsed format object.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function formatRangeCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, range, formatJson] = positionals;
	if (!file || !sheet || !range || !formatJson) {
		throw new AxiError(
			"file, sheet, range and format-json are required",
			"VALIDATION_ERROR",
			["excel-axi format-range <file> <sheet> <range> <format-json>"],
		);
	}

	let format: unknown;
	try {
		format = JSON.parse(formatJson);
	} catch {
		throw new AxiError("format-json is not valid JSON", "VALIDATION_ERROR", [
			'Example: \'{"bold":true,"fontSize":12}\'',
		]);
	}
	if (typeof format !== "object" || format === null || Array.isArray(format)) {
		throw new AxiError("format-json must be a JSON object", "VALIDATION_ERROR");
	}

	return getClient().callTool("excel_format_range", {
		fileAbsolutePath: file,
		sheetName: sheet,
		range,
		...(format as Record<string, unknown>),
	});
}
