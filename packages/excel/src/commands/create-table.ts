/**
 * Why: Defining a table over a range enables structured references and styling; this
 * maps `create-table` to the excel_create_table MCP tool.
 * What: Validates <file> <sheet> <range>, forwards an optional table name to
 * excel_create_table.
 * Test: Mock the client, call createTableCommand(["/tmp/x.xlsx", "Sheet1", "A1:C3", "T"]),
 * assert callTool was invoked with "excel_create_table" and tableName === "T".
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function createTableCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, range, name] = positionals;
	if (!file || !sheet || !range) {
		throw new AxiError("file, sheet and range are required", "VALIDATION_ERROR", [
			"excel-axi create-table <file> <sheet> <range> [name]",
		]);
	}

	const toolArgs: Record<string, unknown> = {
		fileAbsolutePath: file,
		sheetName: sheet,
		range,
	};
	if (name) toolArgs.tableName = name;

	return getClient().callTool("excel_create_table", toolArgs);
}
