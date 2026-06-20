/**
 * Why: Duplicating a sheet is a common workbook edit; this maps `copy-sheet` to the
 * copy_worksheet MCP tool (haris-musa backend).
 * What: Validates <file> <src> <dst> and forwards source/destination sheet names to
 * copy_worksheet using the haris-musa arg names source_sheet / target_sheet.
 * Test: Mock the client, call copySheetCommand(["/tmp/x.xlsx", "A", "B"]), assert
 * callTool was invoked with "copy_worksheet" and { source_sheet: "A", target_sheet: "B" }.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function copySheetCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, src, dst] = positionals;
	if (!file || !src || !dst) {
		throw new AxiError("file, src and dst are required", "VALIDATION_ERROR", [
			"excel-axi copy-sheet <file> <src> <dst>",
		]);
	}

	return getClient().callTool("copy_worksheet", {
		filepath: file,
		source_sheet: src,
		target_sheet: dst,
	});
}
