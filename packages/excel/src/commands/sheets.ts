/**
 * Why: Users need to discover the sheets in a workbook before reading/writing; this
 * maps the `sheets` command to the excel_describe_sheets MCP tool.
 * What: Validates the <file> positional and calls excel_describe_sheets with fileAbsolutePath.
 * Test: Mock the client, call sheetsCommand(["/tmp/x.xlsx"]), assert callTool was
 * invoked with "excel_describe_sheets" and { fileAbsolutePath: "/tmp/x.xlsx" }.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function sheetsCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const file = positionals[0];
	if (!file) {
		throw new AxiError("file path is required", "VALIDATION_ERROR", ["excel-axi sheets <file>"]);
	}

	return getClient().callTool("excel_describe_sheets", {
		fileAbsolutePath: file,
	});
}
