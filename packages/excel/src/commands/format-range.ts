/**
 * Why: Applying styles to a range improves spreadsheet readability; this maps
 * `format-range` to the excel_format_range MCP tool.
 * What: Validates <file> <sheet> <range> <styles-json>, parses the 2D styles array
 * (one style object per cell, matching the grid dimensions of <range>), validates
 * the array shape against the range, and forwards it to excel_format_range.
 * Test: Mock the client, call formatRangeCommand(["/tmp/x.xlsx", "Sheet1", "A1:B2",
 * '[[{"font":{"bold":true}},null],[null,null]]']), assert callTool was invoked with
 * "excel_format_range" and the correct 2D styles array.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

/**
 * Why: The excel_format_range tool requires styles dimensions to match the range grid;
 * this function parses A1:C2-style notation into row and column counts.
 * What: Returns { rows, cols } for the given A1 range string.
 * Test: parseRangeDimensions("A1:C2") → { rows: 2, cols: 3 }.
 */
function parseRangeDimensions(range: string): { rows: number; cols: number } {
	// Match A1:B2 or $A$1:$B$2 style references
	const match = /^\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$/.exec(
		range.trim(),
	);
	if (!match) {
		throw new AxiError(
			`range "${range}" is not a valid A1:Z99 range reference`,
			"VALIDATION_ERROR",
			['Example range: "A1:C2" for 2 rows × 3 cols'],
		);
	}
	const colStart = match[1] ?? "";
	const rowStart = match[2] ?? "";
	const colEnd = match[3] ?? "";
	const rowEnd = match[4] ?? "";

	function colIndex(col: string): number {
		let n = 0;
		for (const ch of col.toUpperCase()) {
			n = n * 26 + (ch.charCodeAt(0) - 64);
		}
		return n;
	}

	const rows = Math.abs(Number(rowEnd) - Number(rowStart)) + 1;
	const cols = Math.abs(colIndex(colEnd) - colIndex(colStart)) + 1;
	return { rows, cols };
}

export async function formatRangeCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, range, stylesJson] = positionals;
	if (!file || !sheet || !range || !stylesJson) {
		throw new AxiError(
			"file, sheet, range and styles-json are required",
			"VALIDATION_ERROR",
			[
				"excel-axi format-range <file> <sheet> <range> <styles-json>",
				"<styles-json> is a 2D JSON array of per-cell style objects matching the range grid.",
				"Example (A1:B2, 2 rows × 2 cols):",
				'  [[{"font":{"bold":true}},{"font":{"italic":true}}],[null,null]]',
				"Use null for cells whose style should not change.",
			],
		);
	}

	let styles: unknown;
	try {
		styles = JSON.parse(stylesJson);
	} catch {
		throw new AxiError("styles-json is not valid JSON", "VALIDATION_ERROR", [
			'Example: \'[[{"font":{"bold":true}},null],[null,null]]\'',
		]);
	}
	if (!Array.isArray(styles)) {
		throw new AxiError(
			"styles-json must be a 2D JSON array (array of row arrays)",
			"VALIDATION_ERROR",
			[
				"Each element of the outer array is a row; each element of a row is either a",
				"style object or null.",
				'Example for A1:B2: [[{"font":{"bold":true}},null],[null,null]]',
			],
		);
	}

	// Validate shape matches range dimensions
	const { rows, cols } = parseRangeDimensions(range);
	if (styles.length !== rows) {
		throw new AxiError(
			`styles-json has ${styles.length} row(s) but range "${range}" spans ${rows} row(s)`,
			"VALIDATION_ERROR",
			[
				`styles-json must be a ${rows}×${cols} 2D array to match range "${range}".`,
				`Outer array length must be ${rows} (one entry per row).`,
			],
		);
	}
	for (let r = 0; r < styles.length; r++) {
		const row = styles[r];
		if (!Array.isArray(row)) {
			throw new AxiError(
				`styles-json row ${r} is not an array`,
				"VALIDATION_ERROR",
				[
					`Each row in styles-json must be an array of ${cols} cell style object(s) or null.`,
				],
			);
		}
		if (row.length !== cols) {
			throw new AxiError(
				`styles-json row ${r} has ${row.length} element(s) but range "${range}" spans ${cols} column(s)`,
				"VALIDATION_ERROR",
				[
					`styles-json must be a ${rows}×${cols} 2D array to match range "${range}".`,
					`Each row must have exactly ${cols} element(s).`,
				],
			);
		}
	}

	return getClient().callTool("excel_format_range", {
		fileAbsolutePath: file,
		sheetName: sheet,
		range,
		styles,
	});
}
