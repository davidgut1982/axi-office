/**
 * Why: Visualizing data with charts is a key Excel feature; this maps `chart` to
 * the create_chart MCP tool.
 * What: Validates <file> <sheet> <data-range> <chart-type> <target-cell> and calls
 * create_chart.
 * Test: Mock the client, call chartCommand(["/tmp/x.xlsx", "Sheet1", "A1:B4", "bar", "E1"]),
 * assert callTool was invoked with "create_chart" and matching args.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

const VALID_CHART_TYPES = ["line", "bar", "pie", "scatter", "area"] as const;

export async function chartCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, sheet, dataRange, chartType, targetCell] = positionals;
	if (!file || !sheet || !dataRange || !chartType || !targetCell) {
		throw new AxiError(
			"file, sheet, data-range, chart-type and target-cell are required",
			"VALIDATION_ERROR",
			[
				"excel-axi chart <file> <sheet> <data-range> <chart-type> <target-cell>",
				"",
				`  chart-type: one of ${VALID_CHART_TYPES.join(", ")}`,
				"  target-cell: top-left corner where chart is placed, e.g. E1",
			]
		);
	}

	if (!VALID_CHART_TYPES.includes(chartType as (typeof VALID_CHART_TYPES)[number])) {
		throw new AxiError(`chart-type "${chartType}" is not supported`, "VALIDATION_ERROR", [
			`Supported types: ${VALID_CHART_TYPES.join(", ")}`,
		]);
	}

	return getClient().callTool("create_chart", {
		filepath: file,
		sheet_name: sheet,
		data_range: dataRange,
		chart_type: chartType,
		target_cell: targetCell,
	});
}
