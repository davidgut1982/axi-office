/**
 * Why: Reading cell data is the most common Excel operation; this maps `read` to
 * the read_data_from_excel MCP tool (haris-musa backend).
 * What: Validates <file> and <sheet>, forwards optional start_cell and end_cell
 * positionals to read_data_from_excel.
 * Test: Mock the client, call readCommand(["/tmp/x.xlsx", "Sheet1", "A1", "D20"]),
 * assert callTool was invoked with "read_data_from_excel" and the matching args.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function readCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, startCell, endCell] = positionals;
	if (!file || !sheet) {
		throw new AxiError("file and sheet are required", "VALIDATION_ERROR", [
			"excel-axi read <file> <sheet> [start-cell] [end-cell]",
			"",
			"  start-cell   Starting cell in A1 notation (default: A1)",
			"  end-cell     Ending cell in A1 notation (optional, auto-expands if omitted)",
		]);
	}

	const toolArgs: Record<string, unknown> = {
		filepath: file,
		sheet_name: sheet,
	};
	if (startCell) toolArgs.start_cell = startCell;
	if (endCell) toolArgs.end_cell = endCell;

	return getClient().callTool("read_data_from_excel", toolArgs);
}
