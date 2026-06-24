/**
 * Why: Charts visualize data trends; this maps `add-chart` to add_chart with full
 * series/category data supplied as JSON arrays.
 * What: Validates <file> <slide-index> <chart-type> and three JSON positionals
 * (categories, series-names, series-values), applies position/size defaults,
 * and calls add_chart via withOpenSave with presentation_id.
 * Test: Mock the client, call addChartCommand with valid args, assert callTool called
 * with "add_chart" and correct categories/series_names/series_values. Also: pass an
 * invalid chart-type and assert AxiError VALIDATION_ERROR.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { call, withOpenSave } from "../session.js";

const VALID_CHART_TYPES = ["column", "bar", "line", "pie"] as const;

export async function addChartCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args);
	const [file, slideIndexRaw, chartType, categoriesJson, seriesNamesJson, seriesValuesJson] =
		positionals;

	if (
		!file ||
		slideIndexRaw === undefined ||
		chartType === undefined ||
		categoriesJson === undefined ||
		seriesNamesJson === undefined ||
		seriesValuesJson === undefined
	) {
		throw new AxiError(
			"file, slide-index, chart-type, categories-json, series-names-json and series-values-json are required",
			"VALIDATION_ERROR",
			[
				"ppt-axi add-chart <file> <slide-index> <chart-type> <categories-json> <series-names-json> <series-values-json>",
				"  [--left=1] [--top=1] [--width=8] [--height=5] [--title T] [--legend-position=right]",
				`  chart-type: one of ${VALID_CHART_TYPES.join(", ")}`,
				'  categories-json:    JSON string array, e.g. \'["Q1","Q2","Q3"]\'',
				"  series-names-json:  JSON string array, e.g. '[\"Revenue\"]'",
				"  series-values-json: JSON array of number arrays, e.g. '[[100,200,300]]'",
			]
		);
	}

	const slideIndex = Number.parseInt(slideIndexRaw, 10);
	if (!Number.isFinite(slideIndex) || slideIndex < 0) {
		throw new AxiError("slide-index must be a non-negative integer", "VALIDATION_ERROR");
	}

	if (!VALID_CHART_TYPES.includes(chartType as (typeof VALID_CHART_TYPES)[number])) {
		throw new AxiError(`chart-type "${chartType}" is not supported`, "VALIDATION_ERROR", [
			`Supported types: ${VALID_CHART_TYPES.join(", ")}`,
		]);
	}

	let categories: unknown;
	try {
		categories = JSON.parse(categoriesJson);
	} catch {
		throw new AxiError("categories-json is not valid JSON", "VALIDATION_ERROR", [
			'Example: \'["Q1","Q2","Q3"]\'',
		]);
	}
	if (!Array.isArray(categories)) {
		throw new AxiError("categories-json must be a JSON array of strings", "VALIDATION_ERROR");
	}

	let seriesNames: unknown;
	try {
		seriesNames = JSON.parse(seriesNamesJson);
	} catch {
		throw new AxiError("series-names-json is not valid JSON", "VALIDATION_ERROR", [
			'Example: \'["Revenue","Cost"]\'',
		]);
	}
	if (!Array.isArray(seriesNames)) {
		throw new AxiError("series-names-json must be a JSON array of strings", "VALIDATION_ERROR");
	}

	let seriesValues: unknown;
	try {
		seriesValues = JSON.parse(seriesValuesJson);
	} catch {
		throw new AxiError("series-values-json is not valid JSON", "VALIDATION_ERROR", [
			"Example: '[[100,200,300],[80,150,250]]'",
		]);
	}
	if (!Array.isArray(seriesValues) || !seriesValues.every((r) => Array.isArray(r))) {
		throw new AxiError(
			"series-values-json must be a 2D JSON array (one array of numbers per series)",
			"VALIDATION_ERROR",
			["Example: '[[100,200,300],[80,150,250]]'"]
		);
	}

	const left = Number.parseFloat(typeof flags.left === "string" ? flags.left : "1");
	const top = Number.parseFloat(typeof flags.top === "string" ? flags.top : "1");
	const width = Number.parseFloat(typeof flags.width === "string" ? flags.width : "8");
	const height = Number.parseFloat(typeof flags.height === "string" ? flags.height : "5");

	return withOpenSave(file, async (client, presentationId) => {
		const toolArgs: Record<string, unknown> = {
			slide_index: slideIndex,
			chart_type: chartType,
			left,
			top,
			width,
			height,
			categories,
			series_names: seriesNames,
			series_values: seriesValues,
			has_legend: true,
			legend_position:
				typeof flags["legend-position"] === "string" ? flags["legend-position"] : "right",
			presentation_id: presentationId,
		};
		if (typeof flags.title === "string") toolArgs.title = flags.title;

		return call(client, "add_chart", toolArgs);
	});
}
