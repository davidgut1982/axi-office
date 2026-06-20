/**
 * Why: Creating a workbook from scratch is the most fundamental operation; this maps
 * `create` to the create_workbook MCP tool (haris-musa backend).
 * What: Validates the <file> positional and calls create_workbook with filepath.
 * Test: Mock the client, call createCommand(["/tmp/x.xlsx"]), assert callTool was
 * invoked with "create_workbook" and { filepath: "/tmp/x.xlsx" }.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function createCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const file = positionals[0];
	if (!file) {
		throw new AxiError("file path is required", "VALIDATION_ERROR", [
			"excel-axi create <file>",
			"",
			"Create a new Excel workbook at <file>.",
			"Requires uv/uvx on PATH (curl -LsSf https://astral.sh/uv/install.sh | sh)",
		]);
	}

	return getClient().callTool("create_workbook", {
		filepath: file,
	});
}
