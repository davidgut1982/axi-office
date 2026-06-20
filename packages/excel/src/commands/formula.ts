/**
 * Why: Applying formulas to cells is essential for computed spreadsheets; this maps
 * `formula` to the apply_formula MCP tool.
 * What: Validates <file> <sheet> <cell> <formula> and calls apply_formula.
 * Test: Mock the client, call formulaCommand(["/tmp/x.xlsx", "Sheet1", "C2", "=SUM(A1:A10)"]),
 * assert callTool was invoked with "apply_formula" and matching args.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function formulaCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, cell, formula] = positionals;
	if (!file || !sheet || !cell || !formula) {
		throw new AxiError("file, sheet, cell and formula are required", "VALIDATION_ERROR", [
			"excel-axi formula <file> <sheet> <cell> <formula>",
			"",
			"  <cell>      Target cell, e.g. C2",
			"  <formula>   Excel formula including = sign, e.g. =SUM(A1:A10)",
		]);
	}

	return getClient().callTool("apply_formula", {
		filepath: file,
		sheet_name: sheet,
		cell,
		formula,
	});
}
