/**
 * Why: Each command must call the correct excel-mcp tool with the correct arguments;
 * these tests lock that mapping so a refactor cannot silently change the wire contract.
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

import { copySheetCommand } from "../src/commands/copy-sheet.js";
import { createTableCommand } from "../src/commands/create-table.js";
import { formatRangeCommand } from "../src/commands/format-range.js";
import { readCommand } from "../src/commands/read.js";
import { sheetsCommand } from "../src/commands/sheets.js";
import { writeCommand } from "../src/commands/write.js";

describe("excel commands", () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCallTool.mockResolvedValue({ ok: true });
	});

	it("sheets calls excel_describe_sheets", async () => {
		await sheetsCommand(["/tmp/x.xlsx"]);
		expect(mockCallTool).toHaveBeenCalledWith("excel_describe_sheets", {
			fileAbsolutePath: "/tmp/x.xlsx",
		});
	});

	it("sheets requires a file", async () => {
		await expect(sheetsCommand([])).rejects.toBeInstanceOf(AxiError);
	});

	it("read calls excel_read_sheet with range", async () => {
		await readCommand(["/tmp/x.xlsx", "Sheet1", "A1:C3"]);
		expect(mockCallTool).toHaveBeenCalledWith("excel_read_sheet", {
			fileAbsolutePath: "/tmp/x.xlsx",
			sheetName: "Sheet1",
			range: "A1:C3",
		});
	});

	it("read passes --formula as showFormula boolean", async () => {
		await readCommand(["/tmp/x.xlsx", "Sheet1", "--formula"]);
		expect(mockCallTool).toHaveBeenCalledWith("excel_read_sheet", {
			fileAbsolutePath: "/tmp/x.xlsx",
			sheetName: "Sheet1",
			showFormula: true,
		});
	});

	it("read passes --style as showStyle boolean", async () => {
		await readCommand(["/tmp/x.xlsx", "Sheet1", "B2:D4", "--style"]);
		expect(mockCallTool).toHaveBeenCalledWith("excel_read_sheet", {
			fileAbsolutePath: "/tmp/x.xlsx",
			sheetName: "Sheet1",
			range: "B2:D4",
			showStyle: true,
		});
	});

	it("read does NOT pass knownPagingRanges", async () => {
		await readCommand(["/tmp/x.xlsx", "Sheet1"]);
		const call = mockCallTool.mock.calls[0]?.[1] as Record<string, unknown>;
		expect(call).not.toHaveProperty("knownPagingRanges");
	});

	it("read requires file and sheet", async () => {
		await expect(readCommand(["/tmp/x.xlsx"])).rejects.toBeInstanceOf(AxiError);
	});

	it("write parses values json and calls excel_write_to_sheet", async () => {
		await writeCommand(["/tmp/x.xlsx", "Sheet1", "A1:B1", '[["a",1]]']);
		expect(mockCallTool).toHaveBeenCalledWith("excel_write_to_sheet", {
			fileAbsolutePath: "/tmp/x.xlsx",
			sheetName: "Sheet1",
			range: "A1:B1",
			values: [["a", 1]],
		});
	});

	it("write rejects invalid json", async () => {
		await expect(
			writeCommand(["/tmp/x.xlsx", "Sheet1", "A1:B1", "not-json"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("create-table calls excel_create_table with name", async () => {
		await createTableCommand(["/tmp/x.xlsx", "Sheet1", "A1:C3", "MyTable"]);
		expect(mockCallTool).toHaveBeenCalledWith("excel_create_table", {
			fileAbsolutePath: "/tmp/x.xlsx",
			sheetName: "Sheet1",
			range: "A1:C3",
			tableName: "MyTable",
		});
	});

	it("copy-sheet calls excel_copy_sheet", async () => {
		await copySheetCommand(["/tmp/x.xlsx", "Src", "Dst"]);
		expect(mockCallTool).toHaveBeenCalledWith("excel_copy_sheet", {
			fileAbsolutePath: "/tmp/x.xlsx",
			srcSheetName: "Src",
			dstSheetName: "Dst",
		});
	});

	it("format-range sends 2D styles array matching the range grid", async () => {
		// A1:B2 → 2 rows × 2 cols → 2D array [[row0col0, row0col1],[row1col0, row1col1]]
		const styles = [[{ font: { bold: true } }, null], [null, null]];
		await formatRangeCommand([
			"/tmp/x.xlsx",
			"Sheet1",
			"A1:B2",
			JSON.stringify(styles),
		]);
		expect(mockCallTool).toHaveBeenCalledWith("excel_format_range", {
			fileAbsolutePath: "/tmp/x.xlsx",
			sheetName: "Sheet1",
			range: "A1:B2",
			styles,
		});
	});

	it("format-range rejects non-array styles json", async () => {
		// A flat object is not the 2D array the tool requires
		await expect(
			formatRangeCommand(["/tmp/x.xlsx", "Sheet1", "A1:B2", '{"bold":true}'])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("format-range rejects styles with wrong row count", async () => {
		// A1:B2 needs 2 rows but we supply 1
		await expect(
			formatRangeCommand(["/tmp/x.xlsx", "Sheet1", "A1:B2", "[[null,null]]"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("format-range rejects styles with wrong column count", async () => {
		// A1:B2 needs 2 cols per row but we supply 3
		await expect(
			formatRangeCommand([
				"/tmp/x.xlsx",
				"Sheet1",
				"A1:B2",
				"[[null,null,null],[null,null,null]]",
			])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("format-range rejects invalid range notation", async () => {
		await expect(
			formatRangeCommand(["/tmp/x.xlsx", "Sheet1", "invalid", "[[null]]"])
		).rejects.toBeInstanceOf(AxiError);
	});
});
