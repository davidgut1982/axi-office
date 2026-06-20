/**
 * Why: Pivot tables summarize large datasets; this maps `pivot` to the
 * create_pivot_table MCP tool.
 * What: Validates <file> <sheet> <data-range> <rows-json> <values-json> and calls
 * create_pivot_table. rows and values are JSON arrays of field name strings.
 * Test: Mock the client, call pivotCommand(["/tmp/x.xlsx", "Sheet1", "A1:C10",
 * '["Name"]', '["Score"]']), assert callTool was invoked with "create_pivot_table"
 * and rows === ["Name"] and values === ["Score"].
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function pivotCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, dataRange, rowsJson, valuesJson] = positionals;
	if (!file || !sheet || !dataRange || !rowsJson || !valuesJson) {
		throw new AxiError(
			"file, sheet, data-range, rows-json and values-json are required",
			"VALIDATION_ERROR",
			[
				"excel-axi pivot <file> <sheet> <data-range> <rows-json> <values-json>",
				"",
				"  <rows-json>    JSON array of field names for row labels, e.g. '[\"Name\"]'",
				"  <values-json>  JSON array of field names to aggregate, e.g. '[\"Score\"]'",
			]
		);
	}

	let rows: unknown;
	let values: unknown;
	try {
		rows = JSON.parse(rowsJson);
	} catch {
		throw new AxiError("rows-json is not valid JSON", "VALIDATION_ERROR", [
			'Example: \'["Name","Region"]\'',
		]);
	}
	try {
		values = JSON.parse(valuesJson);
	} catch {
		throw new AxiError("values-json is not valid JSON", "VALIDATION_ERROR", [
			'Example: \'["Sales","Units"]\'',
		]);
	}
	if (!Array.isArray(rows)) {
		throw new AxiError("rows-json must be a JSON array of field name strings", "VALIDATION_ERROR");
	}
	if (!Array.isArray(values)) {
		throw new AxiError(
			"values-json must be a JSON array of field name strings",
			"VALIDATION_ERROR"
		);
	}

	return getClient().callTool("create_pivot_table", {
		filepath: file,
		sheet_name: sheet,
		data_range: dataRange,
		rows,
		values,
	});
}
