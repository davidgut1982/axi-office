/**
 * Why: Adding sheets to an existing workbook is a common setup step; this maps
 * `create-sheet` to the create_worksheet MCP tool.
 * What: Validates <file> and <sheet> positionals and calls create_worksheet.
 * Test: Mock the client, call createSheetCommand(["/tmp/x.xlsx", "Summary"]), assert
 * callTool was invoked with "create_worksheet" and { filepath: "/tmp/x.xlsx", sheet_name: "Summary" }.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function createSheetCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet] = positionals;
	if (!file || !sheet) {
		throw new AxiError("file and sheet name are required", "VALIDATION_ERROR", [
			"excel-axi create-sheet <file> <sheet>",
		]);
	}

	return getClient().callTool("create_worksheet", {
		filepath: file,
		sheet_name: sheet,
	});
}
