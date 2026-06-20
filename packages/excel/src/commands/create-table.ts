/**
 * Why: Defining a table over a range enables structured references and styling; this
 * maps `table` to the create_table MCP tool (haris-musa backend).
 * What: Validates <file> <sheet> <range>, forwards an optional table name to
 * create_table. Uses `data_range` arg (haris-musa name, not negokaz's `range`).
 * Test: Mock the client, call createTableCommand(["/tmp/x.xlsx", "Sheet1", "A1:C3", "T"]),
 * assert callTool was invoked with "create_table" and data_range === "A1:C3" and table_name === "T".
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function createTableCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, range, name] = positionals;
	if (!file || !sheet || !range) {
		throw new AxiError("file, sheet and range are required", "VALIDATION_ERROR", [
			"excel-axi table <file> <sheet> <range> [name]",
		]);
	}

	const toolArgs: Record<string, unknown> = {
		filepath: file,
		sheet_name: sheet,
		data_range: range,
	};
	if (name) toolArgs.table_name = name;

	return getClient().callTool("create_table", toolArgs);
}
