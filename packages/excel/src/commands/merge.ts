/**
 * Why: Merging cells is a common formatting operation; this maps `merge` to the
 * merge_cells MCP tool.
 * What: Validates <file> <sheet> <start-cell> <end-cell> and calls merge_cells.
 * Test: Mock the client, call mergeCommand(["/tmp/x.xlsx", "Sheet1", "A1", "C1"]),
 * assert callTool was invoked with "merge_cells" and matching args.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function mergeCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, startCell, endCell] = positionals;
	if (!file || !sheet || !startCell || !endCell) {
		throw new AxiError(
			"file, sheet, start-cell and end-cell are required",
			"VALIDATION_ERROR",
			["excel-axi merge <file> <sheet> <start-cell> <end-cell>"]
		);
	}

	return getClient().callTool("merge_cells", {
		filepath: file,
		sheet_name: sheet,
		start_cell: startCell,
		end_cell: endCell,
	});
}
