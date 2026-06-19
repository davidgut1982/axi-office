/**
 * Why: Reading cell values is the most common Excel operation; this maps `read` to the
 * excel_read_sheet MCP tool, with an optional A1 range and a row limit.
 * What: Validates <file> and <sheet>, forwards an optional range and --limit to
 * excel_read_sheet.
 * Test: Mock the client, call readCommand(["/tmp/x.xlsx", "Sheet1", "A1:C3"]), assert
 * callTool was invoked with "excel_read_sheet" and the matching range argument.
 */
import { AxiError, parseFlags, parseLimit } from "@axi-office/core";
import { getClient } from "../client.js";

export async function readCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args);
	const [file, sheet, range] = positionals;
	if (!file || !sheet) {
		throw new AxiError("file and sheet are required", "VALIDATION_ERROR", [
			"excel-axi read <file> <sheet> [range] [--limit N]",
		]);
	}

	const toolArgs: Record<string, unknown> = {
		fileAbsolutePath: file,
		sheetName: sheet,
	};
	if (range) toolArgs.range = range;
	if (flags.limit !== undefined) {
		toolArgs.knownPagingRanges = parseLimit(flags.limit, 100, 10000);
	}

	return getClient().callTool("excel_read_sheet", toolArgs);
}
