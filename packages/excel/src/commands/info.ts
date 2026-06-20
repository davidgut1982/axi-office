/**
 * Why: Users need to inspect a workbook before reading/writing; this maps `info`
 * to the get_workbook_metadata MCP tool.
 * What: Validates the <file> positional and calls get_workbook_metadata with filepath,
 * letting the server use its own default for optional parameters.
 * Test: Mock the client, call infoCommand(["/tmp/x.xlsx"]), assert callTool was
 * invoked with "get_workbook_metadata" and { filepath: "/tmp/x.xlsx" } only (no
 * hardcoded include_ranges).
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function infoCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const file = positionals[0];
	if (!file) {
		throw new AxiError("file path is required", "VALIDATION_ERROR", ["excel-axi info <file>"]);
	}

	return getClient().callTool("get_workbook_metadata", {
		filepath: file,
	});
}
