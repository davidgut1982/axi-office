/**
 * Why: Writing structured values into a sheet is a core operation; this maps `write`
 * to the write_data_to_excel MCP tool (haris-musa backend).
 * What: Validates <file> <sheet> <data-json>, parses the JSON 2D array (array of
 * row arrays), and forwards it to write_data_to_excel. Optional start-cell defaults
 * to A1. Rejects flat arrays (1D) with a clear error pointing at the correct shape.
 * Note: The haris-musa tool takes `data` (List[List]) + `start_cell`, not a range.
 * Test: Mock the client, call writeCommand(["/tmp/x.xlsx", "Sheet1", "[[1,2]]"]),
 * assert callTool was invoked with "write_data_to_excel" and data === [[1, 2]].
 * Also: pass '[1,2]' (1D array) and assert AxiError with VALIDATION_ERROR is thrown.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function writeCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, dataJson, startCell] = positionals;
	if (!file || !sheet || !dataJson) {
		throw new AxiError("file, sheet and data-json are required", "VALIDATION_ERROR", [
			"excel-axi write <file> <sheet> <data-json> [start-cell]",
			"",
			'  <data-json>    JSON array of row arrays, e.g. \'[["Name","Score"],["Alice",90]]\'',
			"  [start-cell]   Cell to start writing at (default: A1)",
		]);
	}

	let data: unknown;
	try {
		data = JSON.parse(dataJson);
	} catch {
		throw new AxiError("data-json is not valid JSON", "VALIDATION_ERROR", [
			'Example: \'[["Name","Score"],["Alice",90]]\'',
		]);
	}
	if (!Array.isArray(data)) {
		throw new AxiError("data-json must be a JSON array of row arrays", "VALIDATION_ERROR");
	}
	if (data.length > 0 && !Array.isArray(data[0])) {
		throw new AxiError("data-json must be a 2D array (array of row arrays)", "VALIDATION_ERROR", [
			'Example: [["Name","Score"],["Alice",90]]',
		]);
	}

	const toolArgs: Record<string, unknown> = {
		filepath: file,
		sheet_name: sheet,
		data,
	};
	if (startCell) toolArgs.start_cell = startCell;

	return getClient().callTool("write_data_to_excel", toolArgs);
}
