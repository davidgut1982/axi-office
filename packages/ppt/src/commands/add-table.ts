/**
 * Why: Tables present structured data clearly in slides; this maps `add-table` to
 * add_table. The backend requires left/top/width/height so CLI defaults are supplied.
 * What: Validates <file> <slide-index> <rows> <cols>, applies position/size defaults,
 * parses optional --data <json-2d> (string[][]) and --header-row flag.
 * Passes presentation_id to add_table.
 * Test: Mock the client, call addTableCommand(["/tmp/x.pptx","0","3","4"]), assert
 * callTool called with "add_table" and rows=3, cols=4 and default dimensions.
 * Also: pass invalid --data and assert AxiError VALIDATION_ERROR.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { call, withOpenSave } from "../session.js";

export async function addTableCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args, ["header-row"]);
	const [file, slideIndexRaw, rowsRaw, colsRaw] = positionals;
	if (!file || slideIndexRaw === undefined || rowsRaw === undefined || colsRaw === undefined) {
		throw new AxiError("file, slide-index, rows and cols are required", "VALIDATION_ERROR", [
			"ppt-axi add-table <file> <slide-index> <rows> <cols> [--left=1] [--top=1] [--width=8] [--height=4]",
			"  [--data <json-2d>] [--header-row]",
			'  --data  2D JSON array of strings, e.g. \'[["Name","Score"],["Alice","90"]]\'',
		]);
	}

	const slideIndex = Number.parseInt(slideIndexRaw, 10);
	if (!Number.isFinite(slideIndex) || slideIndex < 0) {
		throw new AxiError("slide-index must be a non-negative integer", "VALIDATION_ERROR");
	}

	const rows = Number.parseInt(rowsRaw, 10);
	if (!Number.isFinite(rows) || rows < 1) {
		throw new AxiError("rows must be a positive integer", "VALIDATION_ERROR");
	}

	const cols = Number.parseInt(colsRaw, 10);
	if (!Number.isFinite(cols) || cols < 1) {
		throw new AxiError("cols must be a positive integer", "VALIDATION_ERROR");
	}

	const left = Number.parseFloat(typeof flags.left === "string" ? flags.left : "1");
	const top = Number.parseFloat(typeof flags.top === "string" ? flags.top : "1");
	const width = Number.parseFloat(typeof flags.width === "string" ? flags.width : "8");
	const height = Number.parseFloat(typeof flags.height === "string" ? flags.height : "4");

	let data: unknown[][] | undefined;
	if (typeof flags.data === "string") {
		let parsed: unknown;
		try {
			parsed = JSON.parse(flags.data);
		} catch {
			throw new AxiError("--data is not valid JSON", "VALIDATION_ERROR", [
				'Example: \'[["Name","Score"],["Alice","90"]]\'',
			]);
		}
		if (!Array.isArray(parsed) || !parsed.every((r) => Array.isArray(r))) {
			throw new AxiError(
				"--data must be a 2D JSON array (array of row arrays)",
				"VALIDATION_ERROR",
				['Example: \'[["Name","Score"],["Alice","90"]]\'']
			);
		}
		data = parsed as unknown[][];
	}

	return withOpenSave(file, async (client, presentationId) => {
		const toolArgs: Record<string, unknown> = {
			slide_index: slideIndex,
			rows,
			cols,
			left,
			top,
			width,
			height,
			presentation_id: presentationId,
		};
		if (data !== undefined) toolArgs.data = data;
		if (flags["header-row"] === true) toolArgs.header_row = true;

		return call(client, "add_table", toolArgs);
	});
}
