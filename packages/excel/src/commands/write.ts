/**
 * Why: Writing structured values into a sheet is a core operation; this maps `write`
 * to the excel_write_to_sheet MCP tool.
 * What: Validates <file> <sheet> <range> <values-json>, parses the JSON 2D array, and
 * forwards it to excel_write_to_sheet.
 * Test: Mock the client, call writeCommand(["/tmp/x.xlsx", "Sheet1", "A1:B1", "[[1,2]]"]),
 * assert callTool was invoked with "excel_write_to_sheet" and values === [[1,2]].
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function writeCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, range, valuesJson] = positionals;
	if (!file || !sheet || !range || !valuesJson) {
		throw new AxiError("file, sheet, range and values-json are required", "VALIDATION_ERROR", [
			"excel-axi write <file> <sheet> <range> <values-json>",
		]);
	}

	let values: unknown;
	try {
		values = JSON.parse(valuesJson);
	} catch {
		throw new AxiError("values-json is not valid JSON", "VALIDATION_ERROR", [
			'Example: \'[["a","b"],[1,2]]\'',
		]);
	}
	if (!Array.isArray(values)) {
		throw new AxiError("values-json must be a JSON array of rows", "VALIDATION_ERROR");
	}

	return getClient().callTool("excel_write_to_sheet", {
		fileAbsolutePath: file,
		sheetName: sheet,
		range,
		values,
	});
}
