/**
 * Why: Each command must call the correct haris-musa excel-mcp tool with the correct
 * arguments; these tests lock that mapping so a refactor cannot silently change the
 * wire contract.
 * What: Mocks the client factory and asserts the tool name + arguments per command,
 * plus validation errors for missing positionals.
 * Test: This file is the test.
 */
import { AxiError } from "@axi-office/core";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mockCallTool = vi.fn();

vi.mock("../src/client.js", () => ({
	getClient: () => ({ callTool: mockCallTool, close: vi.fn() }),
}));

import { chartCommand } from "../src/commands/chart.js";
import { copySheetCommand } from "../src/commands/copy-sheet.js";
import { createSheetCommand } from "../src/commands/create-sheet.js";
import { createTableCommand } from "../src/commands/create-table.js";
import { createCommand } from "../src/commands/create.js";
import { formatRangeCommand } from "../src/commands/format-range.js";
import { formulaCommand } from "../src/commands/formula.js";
import { infoCommand } from "../src/commands/info.js";
import { mergeCommand } from "../src/commands/merge.js";
import { pivotCommand } from "../src/commands/pivot.js";
import { readCommand } from "../src/commands/read.js";
import { writeCommand } from "../src/commands/write.js";

describe("excel commands (haris-musa backend)", () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCallTool.mockResolvedValue({ ok: true });
	});

	// --- create ---
	it("create calls create_workbook", async () => {
		await createCommand(["/tmp/x.xlsx"]);
		expect(mockCallTool).toHaveBeenCalledWith("create_workbook", {
			filepath: "/tmp/x.xlsx",
		});
	});

	it("create requires a file", async () => {
		await expect(createCommand([])).rejects.toBeInstanceOf(AxiError);
	});

	// --- create-sheet ---
	it("create-sheet calls create_worksheet", async () => {
		await createSheetCommand(["/tmp/x.xlsx", "Summary"]);
		expect(mockCallTool).toHaveBeenCalledWith("create_worksheet", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Summary",
		});
	});

	it("create-sheet requires file and sheet", async () => {
		await expect(createSheetCommand(["/tmp/x.xlsx"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- info ---
	it("info calls get_workbook_metadata without hardcoded include_ranges", async () => {
		await infoCommand(["/tmp/x.xlsx"]);
		expect(mockCallTool).toHaveBeenCalledWith("get_workbook_metadata", {
			filepath: "/tmp/x.xlsx",
		});
	});

	it("info requires a file", async () => {
		await expect(infoCommand([])).rejects.toBeInstanceOf(AxiError);
	});

	// --- read ---
	it("read calls read_data_from_excel with start and end cell", async () => {
		await readCommand(["/tmp/x.xlsx", "Sheet1", "A1", "D20"]);
		expect(mockCallTool).toHaveBeenCalledWith("read_data_from_excel", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			start_cell: "A1",
			end_cell: "D20",
		});
	});

	it("read calls read_data_from_excel without optional cells", async () => {
		await readCommand(["/tmp/x.xlsx", "Sheet1"]);
		expect(mockCallTool).toHaveBeenCalledWith("read_data_from_excel", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
		});
	});

	it("read requires file and sheet", async () => {
		await expect(readCommand(["/tmp/x.xlsx"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- write ---
	it("write calls write_data_to_excel with data array", async () => {
		await writeCommand(["/tmp/x.xlsx", "Sheet1", '[["Name","Score"],["Alice",90]]']);
		expect(mockCallTool).toHaveBeenCalledWith("write_data_to_excel", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			data: [
				["Name", "Score"],
				["Alice", 90],
			],
		});
	});

	it("write passes optional start_cell", async () => {
		await writeCommand(["/tmp/x.xlsx", "Sheet1", "[[1,2]]", "B2"]);
		expect(mockCallTool).toHaveBeenCalledWith("write_data_to_excel", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			data: [[1, 2]],
			start_cell: "B2",
		});
	});

	it("write rejects invalid json", async () => {
		await expect(writeCommand(["/tmp/x.xlsx", "Sheet1", "not-json"])).rejects.toBeInstanceOf(
			AxiError
		);
	});

	it("write rejects non-array json", async () => {
		await expect(writeCommand(["/tmp/x.xlsx", "Sheet1", '{"key":"val"}'])).rejects.toBeInstanceOf(
			AxiError
		);
	});

	it("write rejects a flat (1D) array", async () => {
		await expect(
			writeCommand(["/tmp/x.xlsx", "Sheet1", '["Name","Score"]'])
		).rejects.toBeInstanceOf(AxiError);
	});

	// --- formula ---
	it("formula calls apply_formula", async () => {
		await formulaCommand(["/tmp/x.xlsx", "Sheet1", "C2", "=SUM(A1:A10)"]);
		expect(mockCallTool).toHaveBeenCalledWith("apply_formula", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			cell: "C2",
			formula: "=SUM(A1:A10)",
		});
	});

	it("formula requires all four args", async () => {
		await expect(formulaCommand(["/tmp/x.xlsx", "Sheet1", "C2"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- format-range ---
	it("format-range spreads flat style object into tool args with --end flag", async () => {
		await formatRangeCommand([
			"/tmp/x.xlsx",
			"Sheet1",
			"A1",
			'{"bold":true,"font_size":14}',
			"--end",
			"B1",
		]);
		expect(mockCallTool).toHaveBeenCalledWith("format_range", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			start_cell: "A1",
			end_cell: "B1",
			bold: true,
			font_size: 14,
		});
	});

	it("format-range without --end flag omits end_cell", async () => {
		await formatRangeCommand(["/tmp/x.xlsx", "Sheet1", "A1", '{"italic":true}']);
		const call = mockCallTool.mock.calls[0]?.[1] as Record<string, unknown>;
		expect(call).not.toHaveProperty("end_cell");
		expect(call).toHaveProperty("italic", true);
	});

	it("format-range rejects a 2D array (negokaz shape)", async () => {
		await expect(
			formatRangeCommand(["/tmp/x.xlsx", "Sheet1", "A1", "[[null,null]]"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("format-range rejects invalid json", async () => {
		await expect(
			formatRangeCommand(["/tmp/x.xlsx", "Sheet1", "A1", "not-json"])
		).rejects.toBeInstanceOf(AxiError);
	});

	// --- merge ---
	it("merge calls merge_cells", async () => {
		await mergeCommand(["/tmp/x.xlsx", "Sheet1", "A1", "C1"]);
		expect(mockCallTool).toHaveBeenCalledWith("merge_cells", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			start_cell: "A1",
			end_cell: "C1",
		});
	});

	it("merge requires all four args", async () => {
		await expect(mergeCommand(["/tmp/x.xlsx", "Sheet1", "A1"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- table ---
	it("table calls create_table with data_range", async () => {
		await createTableCommand(["/tmp/x.xlsx", "Sheet1", "A1:C3", "MyTable"]);
		expect(mockCallTool).toHaveBeenCalledWith("create_table", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			data_range: "A1:C3",
			table_name: "MyTable",
		});
	});

	it("table omits table_name when not provided", async () => {
		await createTableCommand(["/tmp/x.xlsx", "Sheet1", "A1:C3"]);
		const call = mockCallTool.mock.calls[0]?.[1] as Record<string, unknown>;
		expect(call).not.toHaveProperty("table_name");
		expect(call).toHaveProperty("data_range", "A1:C3");
	});

	// --- chart ---
	it("chart calls create_chart", async () => {
		await chartCommand(["/tmp/x.xlsx", "Sheet1", "A1:B4", "bar", "E1"]);
		expect(mockCallTool).toHaveBeenCalledWith("create_chart", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			data_range: "A1:B4",
			chart_type: "bar",
			target_cell: "E1",
		});
	});

	it("chart rejects invalid chart type", async () => {
		await expect(
			chartCommand(["/tmp/x.xlsx", "Sheet1", "A1:B4", "donut", "E1"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("chart requires all five args", async () => {
		await expect(chartCommand(["/tmp/x.xlsx", "Sheet1", "A1:B4", "bar"])).rejects.toBeInstanceOf(
			AxiError
		);
	});

	// --- pivot ---
	it("pivot calls create_pivot_table with rows, values, and default agg_func=sum", async () => {
		await pivotCommand(["/tmp/x.xlsx", "Sheet1", "A1:C10", '["Name"]', '["Score"]']);
		expect(mockCallTool).toHaveBeenCalledWith("create_pivot_table", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			data_range: "A1:C10",
			rows: ["Name"],
			values: ["Score"],
			agg_func: "sum",
		});
	});

	it("pivot passes explicit --agg flag", async () => {
		await pivotCommand([
			"/tmp/x.xlsx",
			"Sheet1",
			"A1:C10",
			'["Name"]',
			'["Score"]',
			"--agg",
			"average",
		]);
		expect(mockCallTool).toHaveBeenCalledWith("create_pivot_table", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			data_range: "A1:C10",
			rows: ["Name"],
			values: ["Score"],
			agg_func: "average",
		});
	});

	it("pivot rejects invalid --agg value", async () => {
		await expect(
			pivotCommand(["/tmp/x.xlsx", "Sheet1", "A1:C10", '["Name"]', '["Score"]', "--agg", "mean"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("pivot rejects invalid rows-json", async () => {
		await expect(
			pivotCommand(["/tmp/x.xlsx", "Sheet1", "A1:C10", "not-json", '["Score"]'])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("pivot requires all five args", async () => {
		await expect(
			pivotCommand(["/tmp/x.xlsx", "Sheet1", "A1:C10", '["Name"]'])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("pivot passes --columns when provided", async () => {
		await pivotCommand([
			"/tmp/x.xlsx",
			"Sheet1",
			"A1:C10",
			'["Name"]',
			'["Score"]',
			"--columns",
			'["Region"]',
		]);
		expect(mockCallTool).toHaveBeenCalledWith("create_pivot_table", {
			filepath: "/tmp/x.xlsx",
			sheet_name: "Sheet1",
			data_range: "A1:C10",
			rows: ["Name"],
			values: ["Score"],
			agg_func: "sum",
			columns: ["Region"],
		});
	});

	it("pivot omits columns from tool args when --columns is not given", async () => {
		await pivotCommand(["/tmp/x.xlsx", "Sheet1", "A1:C10", '["Name"]', '["Score"]']);
		const call = mockCallTool.mock.calls[0]?.[1] as Record<string, unknown>;
		expect(call).not.toHaveProperty("columns");
	});

	it("pivot rejects invalid --columns json", async () => {
		await expect(
			pivotCommand([
				"/tmp/x.xlsx",
				"Sheet1",
				"A1:C10",
				'["Name"]',
				'["Score"]',
				"--columns",
				"not-json",
			])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("pivot rejects non-array --columns value", async () => {
		await expect(
			pivotCommand([
				"/tmp/x.xlsx",
				"Sheet1",
				"A1:C10",
				'["Name"]',
				'["Score"]',
				"--columns",
				'"Region"',
			])
		).rejects.toBeInstanceOf(AxiError);
	});

	// --- copy-sheet ---
	it("copy-sheet calls copy_worksheet with source_sheet / target_sheet", async () => {
		await copySheetCommand(["/tmp/x.xlsx", "Src", "Dst"]);
		expect(mockCallTool).toHaveBeenCalledWith("copy_worksheet", {
			filepath: "/tmp/x.xlsx",
			source_sheet: "Src",
			target_sheet: "Dst",
		});
	});

	it("copy-sheet requires all three args", async () => {
		await expect(copySheetCommand(["/tmp/x.xlsx", "Src"])).rejects.toBeInstanceOf(AxiError);
	});
});
