/**
 * Why: Duplicating a sheet is a common workbook edit; this maps `copy-sheet` to the
 * excel_copy_sheet MCP tool.
 * What: Validates <file> <src> <dst> and forwards source/destination sheet names to
 * excel_copy_sheet.
 * Test: Mock the client, call copySheetCommand(["/tmp/x.xlsx", "A", "B"]), assert
 * callTool was invoked with "excel_copy_sheet" and { srcSheetName: "A", dstSheetName: "B" }.
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

	return getClient().callTool("excel_copy_sheet", {
		fileAbsolutePath: file,
		srcSheetName: src,
		dstSheetName: dst,
	});
}
