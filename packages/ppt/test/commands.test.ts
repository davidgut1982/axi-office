/**
 * Why: Each command must call the correct GongRzhe PowerPoint MCP tools in the correct
 * order with the correct arguments; these tests lock that mapping so a refactor cannot
 * silently change the wire contract.
 * What: Mocks the client factory and asserts the tool name + arguments per command,
 * plus the open→op→save sequencing for mutating commands, and that read-only commands
 * do NOT call save_presentation. Validates error conditions for malformed inputs.
 * Test: This file is the test.
 */
import { AxiError } from "@axi-office/core";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mockCallTool = vi.fn();

vi.mock("../src/client.js", () => ({
	getClient: () => ({ callTool: mockCallTool, close: vi.fn() }),
}));

import { addChartCommand } from "../src/commands/add-chart.js";
import { addImageCommand } from "../src/commands/add-image.js";
import { addShapeCommand } from "../src/commands/add-shape.js";
import { addSlideCommand } from "../src/commands/add-slide.js";
import { addTableCommand } from "../src/commands/add-table.js";
import { addTextCommand } from "../src/commands/add-text.js";
import { autoGenerateCommand } from "../src/commands/auto-generate.js";
import { bulletsCommand } from "../src/commands/bullets.js";
import { createCommand } from "../src/commands/create.js";
import { fromTemplateCommand } from "../src/commands/from-template.js";
import { infoCommand } from "../src/commands/info.js";
import { readCommand } from "../src/commands/read.js";
import { readSlideCommand } from "../src/commands/read-slide.js";
import { setPlaceholderCommand } from "../src/commands/set-placeholder.js";
import { setPropsCommand } from "../src/commands/set-props.js";
import { slideInfoCommand } from "../src/commands/slide-info.js";
import { templatesCommand } from "../src/commands/templates.js";

describe("ppt commands (GongRzhe backend)", () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCallTool.mockImplementation(async (tool: string) => {
			if (
				tool === "create_presentation" ||
				tool === "open_presentation" ||
				tool === "create_presentation_from_template"
			) {
				return { presentation_id: "presentation_1" };
			}
			return { ok: true };
		});
	});

	// helper: return all tool names called in order
	function calledTools(): string[] {
		return mockCallTool.mock.calls.map((c) => c[0] as string);
	}

	// --- create ---
	it("create calls create_presentation then save_presentation", async () => {
		await createCommand(["/tmp/x.pptx"]);
		expect(calledTools()).toEqual(["create_presentation", "save_presentation"]);
		expect(mockCallTool.mock.calls[1]?.[1]).toMatchObject({
			file_path: "/tmp/x.pptx",
			presentation_id: "presentation_1",
		});
	});

	it("create requires a file", async () => {
		await expect(createCommand([])).rejects.toBeInstanceOf(AxiError);
	});

	// --- info ---
	it("info calls open_presentation then get_presentation_info (no save)", async () => {
		await infoCommand(["/tmp/x.pptx"]);
		expect(calledTools()).toEqual(["open_presentation", "get_presentation_info"]);
		expect(calledTools()).not.toContain("save_presentation");
	});

	it("info requires a file", async () => {
		await expect(infoCommand([])).rejects.toBeInstanceOf(AxiError);
	});

	// --- set-props ---
	it("set-props calls open→set_core_properties→save with provided flags only", async () => {
		await setPropsCommand(["/tmp/x.pptx", "--title", "My Deck"]);
		expect(calledTools()).toEqual([
			"open_presentation",
			"set_core_properties",
			"save_presentation",
		]);
		expect(mockCallTool.mock.calls[1]?.[1]).toEqual({
			title: "My Deck",
			presentation_id: "presentation_1",
		});
	});

	it("set-props passes multiple props", async () => {
		await setPropsCommand(["/tmp/x.pptx", "--title", "T", "--author", "A"]);
		expect(mockCallTool.mock.calls[1]?.[1]).toEqual({
			title: "T",
			author: "A",
			presentation_id: "presentation_1",
		});
	});

	it("set-props requires file", async () => {
		await expect(setPropsCommand([])).rejects.toBeInstanceOf(AxiError);
	});

	it("set-props requires at least one property flag", async () => {
		await expect(setPropsCommand(["/tmp/x.pptx"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- from-template ---
	it("from-template calls create_presentation_from_template then save_presentation", async () => {
		await fromTemplateCommand(["/tmp/out.pptx", "/tmp/tmpl.pptx"]);
		expect(calledTools()).toEqual([
			"create_presentation_from_template",
			"save_presentation",
		]);
		expect(mockCallTool.mock.calls[0]?.[1]).toMatchObject({ template_path: "/tmp/tmpl.pptx" });
		expect(mockCallTool.mock.calls[1]?.[1]).toMatchObject({
			file_path: "/tmp/out.pptx",
			presentation_id: "presentation_1",
		});
	});

	it("from-template requires file and template-path", async () => {
		await expect(fromTemplateCommand(["/tmp/out.pptx"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- add-slide ---
	it("add-slide calls open→add_slide→save with default layout_index=1", async () => {
		await addSlideCommand(["/tmp/x.pptx"]);
		expect(calledTools()).toEqual(["open_presentation", "add_slide", "save_presentation"]);
		expect(mockCallTool.mock.calls[1]?.[1]).toMatchObject({
			layout_index: 1,
			presentation_id: "presentation_1",
		});
		expect(mockCallTool.mock.calls[2]?.[1]).toMatchObject({
			file_path: "/tmp/x.pptx",
			presentation_id: "presentation_1",
		});
	});

	it("add-slide passes --title and --color-scheme when provided", async () => {
		await addSlideCommand(["/tmp/x.pptx", "--title", "Agenda", "--color-scheme", "modern_blue"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).toMatchObject({ title: "Agenda", color_scheme: "modern_blue" });
	});

	it("add-slide uses --layout value", async () => {
		await addSlideCommand(["/tmp/x.pptx", "--layout", "3"]);
		expect(mockCallTool.mock.calls[1]?.[1]).toMatchObject({ layout_index: 3 });
	});

	it("add-slide requires a file", async () => {
		await expect(addSlideCommand([])).rejects.toBeInstanceOf(AxiError);
	});

	it("add-slide rejects invalid --layout", async () => {
		await expect(addSlideCommand(["/tmp/x.pptx", "--layout", "abc"])).rejects.toBeInstanceOf(
			AxiError
		);
	});

	// --- slide-info ---
	it("slide-info calls open_presentation then get_slide_info (no save)", async () => {
		await slideInfoCommand(["/tmp/x.pptx", "0"]);
		expect(calledTools()).toEqual(["open_presentation", "get_slide_info"]);
		expect(mockCallTool.mock.calls[1]?.[1]).toEqual({
			slide_index: 0,
			presentation_id: "presentation_1",
		});
		expect(calledTools()).not.toContain("save_presentation");
	});

	it("slide-info requires file and slide-index", async () => {
		await expect(slideInfoCommand(["/tmp/x.pptx"])).rejects.toBeInstanceOf(AxiError);
	});

	it("slide-info rejects negative slide-index", async () => {
		await expect(slideInfoCommand(["/tmp/x.pptx", "-1"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- read ---
	it("read calls open_presentation then extract_presentation_text (no save)", async () => {
		await readCommand(["/tmp/x.pptx"]);
		expect(calledTools()).toEqual(["open_presentation", "extract_presentation_text"]);
		expect(mockCallTool.mock.calls[1]?.[1]).toEqual({
			include_slide_info: true,
			presentation_id: "presentation_1",
		});
		expect(calledTools()).not.toContain("save_presentation");
	});

	it("read requires a file", async () => {
		await expect(readCommand([])).rejects.toBeInstanceOf(AxiError);
	});

	// --- read-slide ---
	it("read-slide calls open_presentation then extract_slide_text (no save)", async () => {
		await readSlideCommand(["/tmp/x.pptx", "2"]);
		expect(calledTools()).toEqual(["open_presentation", "extract_slide_text"]);
		expect(mockCallTool.mock.calls[1]?.[1]).toEqual({
			slide_index: 2,
			presentation_id: "presentation_1",
		});
		expect(calledTools()).not.toContain("save_presentation");
	});

	it("read-slide requires file and slide-index", async () => {
		await expect(readSlideCommand(["/tmp/x.pptx"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- set-placeholder ---
	it("set-placeholder calls open→populate_placeholder→save", async () => {
		await setPlaceholderCommand(["/tmp/x.pptx", "0", "0", "Hello"]);
		expect(calledTools()).toEqual([
			"open_presentation",
			"populate_placeholder",
			"save_presentation",
		]);
		expect(mockCallTool.mock.calls[1]?.[1]).toEqual({
			slide_index: 0,
			placeholder_idx: 0,
			text: "Hello",
			presentation_id: "presentation_1",
		});
	});

	it("set-placeholder requires all four args", async () => {
		await expect(
			setPlaceholderCommand(["/tmp/x.pptx", "0", "0"])
		).rejects.toBeInstanceOf(AxiError);
	});

	// --- bullets ---
	it("bullets calls open→add_bullet_points→save with parsed array", async () => {
		await bulletsCommand(["/tmp/x.pptx", "0", "1", '["A","B","C"]']);
		expect(calledTools()).toEqual([
			"open_presentation",
			"add_bullet_points",
			"save_presentation",
		]);
		expect(mockCallTool.mock.calls[1]?.[1]).toEqual({
			slide_index: 0,
			placeholder_idx: 1,
			bullet_points: ["A", "B", "C"],
			presentation_id: "presentation_1",
		});
	});

	it("bullets rejects invalid JSON", async () => {
		await expect(
			bulletsCommand(["/tmp/x.pptx", "0", "1", "not-json"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("bullets rejects non-array JSON", async () => {
		await expect(
			bulletsCommand(["/tmp/x.pptx", "0", "1", '"a string"'])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("bullets requires all four args", async () => {
		await expect(bulletsCommand(["/tmp/x.pptx", "0", "1"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- add-text ---
	it("add-text calls open→manage_text→save with operation=add and defaults", async () => {
		await addTextCommand(["/tmp/x.pptx", "0", "Hello World"]);
		expect(calledTools()).toEqual(["open_presentation", "manage_text", "save_presentation"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).toMatchObject({
			operation: "add",
			slide_index: 0,
			text: "Hello World",
			left: 1,
			top: 1,
			width: 4,
			height: 2,
			presentation_id: "presentation_1",
		});
		expect(mockCallTool.mock.calls[2]?.[1]).toMatchObject({
			file_path: "/tmp/x.pptx",
			presentation_id: "presentation_1",
		});
	});

	it("add-text passes position flags", async () => {
		await addTextCommand(["/tmp/x.pptx", "1", "Hi", "--left", "2", "--top", "3"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).toMatchObject({ left: 2, top: 3 });
	});

	it("add-text parses --color R,G,B into array", async () => {
		await addTextCommand(["/tmp/x.pptx", "0", "Hi", "--color", "255,0,128"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs.color).toEqual([255, 0, 128]);
	});

	it("add-text rejects invalid --color", async () => {
		await expect(
			addTextCommand(["/tmp/x.pptx", "0", "Hi", "--color", "notacolor"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-text requires file, slide-index and text", async () => {
		await expect(addTextCommand(["/tmp/x.pptx", "0"])).rejects.toBeInstanceOf(AxiError);
	});

	it("add-text rejects non-numeric --font-size", async () => {
		await expect(
			addTextCommand(["/tmp/x.pptx", "0", "Hi", "--font-size", "abc"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-text rejects non-numeric --left", async () => {
		await expect(
			addTextCommand(["/tmp/x.pptx", "0", "Hi", "--left", "foo"])
		).rejects.toBeInstanceOf(AxiError);
	});

	// --- add-image ---
	it("add-image calls open→manage_image→save with operation=add source_type=file", async () => {
		await addImageCommand(["/tmp/x.pptx", "0", "/img.png"]);
		expect(calledTools()).toEqual(["open_presentation", "manage_image", "save_presentation"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).toMatchObject({
			operation: "add",
			slide_index: 0,
			image_source: "/img.png",
			source_type: "file",
			left: 1,
			top: 1,
			presentation_id: "presentation_1",
		});
		expect(toolArgs).not.toHaveProperty("width");
		expect(toolArgs).not.toHaveProperty("height");
	});

	it("add-image passes --width and --height when provided", async () => {
		await addImageCommand(["/tmp/x.pptx", "0", "/img.png", "--width", "4", "--height", "3"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).toMatchObject({ width: 4, height: 3 });
	});

	it("add-image rejects non-numeric --width", async () => {
		await expect(
			addImageCommand(["/tmp/x.pptx", "0", "/img.png", "--width", "foo"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-image requires file, slide-index and image-path", async () => {
		await expect(addImageCommand(["/tmp/x.pptx", "0"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- add-table ---
	it("add-table calls open→add_table→save with defaults", async () => {
		await addTableCommand(["/tmp/x.pptx", "0", "3", "4"]);
		expect(calledTools()).toEqual(["open_presentation", "add_table", "save_presentation"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).toMatchObject({
			slide_index: 0,
			rows: 3,
			cols: 4,
			left: 1,
			top: 1,
			width: 8,
			height: 4,
			presentation_id: "presentation_1",
		});
		expect(toolArgs).not.toHaveProperty("data");
		expect(toolArgs).not.toHaveProperty("header_row");
		expect(mockCallTool.mock.calls[2]?.[1]).toMatchObject({
			file_path: "/tmp/x.pptx",
			presentation_id: "presentation_1",
		});
	});

	it("add-table passes --data and --header-row when provided", async () => {
		await addTableCommand([
			"/tmp/x.pptx",
			"1",
			"2",
			"2",
			"--data",
			'[["Name","Score"],["Alice","90"]]',
			"--header-row",
		]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs.data).toEqual([["Name", "Score"], ["Alice", "90"]]);
		expect(toolArgs.header_row).toBe(true);
	});

	it("add-table rejects invalid --data JSON", async () => {
		await expect(
			addTableCommand(["/tmp/x.pptx", "0", "2", "2", "--data", "not-json"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-table rejects flat (1D) --data", async () => {
		await expect(
			addTableCommand(["/tmp/x.pptx", "0", "2", "2", "--data", '["a","b"]'])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-table rejects mixed --data where later rows are not arrays", async () => {
		await expect(
			addTableCommand(["/tmp/x.pptx", "0", "2", "2", "--data", '[["a","b"],"bad"]'])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-table rejects invalid rows value", async () => {
		await expect(
			addTableCommand(["/tmp/x.pptx", "0", "0", "3"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-table rejects non-numeric --width", async () => {
		await expect(
			addTableCommand(["/tmp/x.pptx", "0", "3", "4", "--width", "foo"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-table requires file, slide-index, rows, cols", async () => {
		await expect(addTableCommand(["/tmp/x.pptx", "0", "3"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- add-shape ---
	it("add-shape calls open→add_shape→save with defaults", async () => {
		await addShapeCommand(["/tmp/x.pptx", "0", "rectangle"]);
		expect(calledTools()).toEqual(["open_presentation", "add_shape", "save_presentation"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).toMatchObject({
			slide_index: 0,
			shape_type: "rectangle",
			left: 1,
			top: 1,
			width: 2,
			height: 2,
			presentation_id: "presentation_1",
		});
	});

	it("add-shape passes --fill-color as [R,G,B] array", async () => {
		await addShapeCommand(["/tmp/x.pptx", "0", "oval", "--fill-color", "0,112,192"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs.fill_color).toEqual([0, 112, 192]);
	});

	it("add-shape passes --text when provided", async () => {
		await addShapeCommand(["/tmp/x.pptx", "0", "rectangle", "--text", "Key Point"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs.text).toBe("Key Point");
	});

	it("add-shape rejects invalid --fill-color", async () => {
		await expect(
			addShapeCommand(["/tmp/x.pptx", "0", "rectangle", "--fill-color", "notacolor"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-shape rejects non-numeric --left", async () => {
		await expect(
			addShapeCommand(["/tmp/x.pptx", "0", "rectangle", "--left", "foo"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-shape requires file, slide-index and shape-type", async () => {
		await expect(addShapeCommand(["/tmp/x.pptx", "0"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- add-chart ---
	it("add-chart calls open→add_chart→save with correct args", async () => {
		await addChartCommand([
			"/tmp/x.pptx",
			"2",
			"column",
			'["Q1","Q2","Q3"]',
			'["Revenue"]',
			'[[100,200,300]]',
		]);
		expect(calledTools()).toEqual(["open_presentation", "add_chart", "save_presentation"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).toMatchObject({
			slide_index: 2,
			chart_type: "column",
			categories: ["Q1", "Q2", "Q3"],
			series_names: ["Revenue"],
			series_values: [[100, 200, 300]],
			has_legend: true,
			legend_position: "right",
			presentation_id: "presentation_1",
		});
	});

	it("add-chart passes --title when provided", async () => {
		await addChartCommand([
			"/tmp/x.pptx",
			"0",
			"bar",
			'["A"]',
			'["S"]',
			'[[1]]',
			"--title",
			"My Chart",
		]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs.title).toBe("My Chart");
	});

	it("add-chart omits title when not provided", async () => {
		await addChartCommand(["/tmp/x.pptx", "0", "line", '["A"]', '["S"]', '[[1]]']);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).not.toHaveProperty("title");
	});

	it("add-chart rejects invalid chart-type", async () => {
		await expect(
			addChartCommand(["/tmp/x.pptx", "0", "donut", '["A"]', '["S"]', '[[1]]'])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-chart rejects invalid categories-json", async () => {
		await expect(
			addChartCommand(["/tmp/x.pptx", "0", "bar", "not-json", '["S"]', '[[1]]'])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-chart rejects flat (1D) series-values", async () => {
		await expect(
			addChartCommand(["/tmp/x.pptx", "0", "bar", '["A"]', '["S"]', '[1,2,3]'])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-chart rejects mixed series-values where a later element is not an array", async () => {
		await expect(
			addChartCommand(["/tmp/x.pptx", "0", "bar", '["A"]', '["S"]', '[[1,2],42]'])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-chart rejects non-numeric --width", async () => {
		await expect(
			addChartCommand(["/tmp/x.pptx", "0", "bar", '["A"]', '["S"]', '[[1]]', "--width", "foo"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("add-chart requires all six positionals", async () => {
		await expect(
			addChartCommand(["/tmp/x.pptx", "0", "column", '["A"]', '["S"]'])
		).rejects.toBeInstanceOf(AxiError);
	});

	// --- templates ---
	it("templates calls list_slide_templates directly (no open or save)", async () => {
		await templatesCommand([]);
		expect(calledTools()).toEqual(["list_slide_templates"]);
	});

	// --- auto-generate ---
	it("auto-generate calls create_presentation→auto_generate_presentation→save_presentation", async () => {
		await autoGenerateCommand(["/tmp/x.pptx", "AI trends"]);
		expect(calledTools()).toEqual([
			"create_presentation",
			"auto_generate_presentation",
			"save_presentation",
		]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).toMatchObject({
			topic: "AI trends",
			slide_count: 5,
			presentation_type: "business",
			color_scheme: "modern_blue",
			include_charts: true,
			presentation_id: "presentation_1",
		});
		expect(mockCallTool.mock.calls[2]?.[1]).toMatchObject({
			file_path: "/tmp/x.pptx",
			presentation_id: "presentation_1",
		});
	});

	it("auto-generate passes --slides, --type, --color-scheme", async () => {
		await autoGenerateCommand([
			"/tmp/x.pptx",
			"topic",
			"--slides",
			"8",
			"--type",
			"academic",
			"--color-scheme",
			"elegant_green",
		]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs).toMatchObject({
			slide_count: 8,
			presentation_type: "academic",
			color_scheme: "elegant_green",
		});
	});

	it("auto-generate --no-charts sets include_charts=false", async () => {
		await autoGenerateCommand(["/tmp/x.pptx", "topic", "--no-charts"]);
		const toolArgs = mockCallTool.mock.calls[1]?.[1] as Record<string, unknown>;
		expect(toolArgs.include_charts).toBe(false);
	});

	it("auto-generate rejects invalid --type", async () => {
		await expect(
			autoGenerateCommand(["/tmp/x.pptx", "topic", "--type", "unknown"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("auto-generate rejects invalid --color-scheme", async () => {
		await expect(
			autoGenerateCommand(["/tmp/x.pptx", "topic", "--color-scheme", "rainbow"])
		).rejects.toBeInstanceOf(AxiError);
	});

	it("auto-generate requires file and topic", async () => {
		await expect(autoGenerateCommand(["/tmp/x.pptx"])).rejects.toBeInstanceOf(AxiError);
	});

	// --- BUG 2: backend error surfacing ---
	it("surfaces backend error when callTool returns { error: '...' }", async () => {
		mockCallTool.mockImplementation(async (tool: string) => {
			if (tool === "open_presentation") return { presentation_id: "presentation_1" };
			if (tool === "get_presentation_info")
				return { error: "No presentation is currently loaded" };
			return { ok: true };
		});
		await expect(infoCommand(["/tmp/x.pptx"])).rejects.toBeInstanceOf(AxiError);
	});
});
