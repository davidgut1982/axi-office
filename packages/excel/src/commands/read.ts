/**
 * Why: Reading cell values is the most common Excel operation; this maps `read` to the
 * excel_read_sheet MCP tool, with an optional A1 range and optional formula/style flags.
 * What: Validates <file> and <sheet>, forwards an optional range, --formula, and --style
 * boolean flags to excel_read_sheet.
 * Test: Mock the client, call readCommand(["/tmp/x.xlsx", "Sheet1", "A1:C3"]), assert
 * callTool was invoked with "excel_read_sheet" and the matching range argument.
 *
 * Note: Cell paging is controlled server-side by the EXCEL_MCP_PAGING_CELLS_LIMIT env var
 * (default 2000) set on the excel-mcp-server process. There is no client-side paging flag.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function readCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args, ["formula", "style"]);
	const [file, sheet, range] = positionals;
	if (!file || !sheet) {
		throw new AxiError("file and sheet are required", "VALIDATION_ERROR", [
			"excel-axi read <file> <sheet> [range] [--formula] [--style]",
			"",
			"  --formula    Show cell formulas instead of computed values",
			"  --style      Include style information for each cell",
			"",
			"Tip: set EXCEL_MCP_PAGING_CELLS_LIMIT=N on the server process to control",
			"     how many cells are returned per page (default: 2000).",
		]);
	}

	const toolArgs: Record<string, unknown> = {
		fileAbsolutePath: file,
		sheetName: sheet,
	};
	if (range) toolArgs.range = range;
	if (flags.formula === true) toolArgs.showFormula = true;
	if (flags.style === true) toolArgs.showStyle = true;

	return getClient().callTool("excel_read_sheet", toolArgs);
}
