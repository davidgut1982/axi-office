/**
 * Why: runAxiCli needs a top-level help string and per-command help lookup; keeping
 * them in one module avoids scattering usage strings across command files.
 * What: Exports TOP_LEVEL_HELP and a COMMAND_HELP record keyed by command name.
 * Test: Assert TOP_LEVEL_HELP contains "excel-axi" and COMMAND_HELP.create is a
 * non-empty string.
 */

export const TOP_LEVEL_HELP = `excel-axi — AXI CLI for Excel (via haris-musa/excel-mcp-server, run with uvx)

Requires: uv/uvx on PATH. Install: curl -LsSf https://astral.sh/uv/install.sh | sh

Usage: excel-axi <command> [args] [flags]

Commands:
  create        Create a new Excel workbook from scratch
  create-sheet  Add a new worksheet to a workbook
  info          Show workbook metadata (sheets, size)
  read          Read data from a worksheet
  write         Write a 2D JSON array to a worksheet
  formula       Apply an Excel formula to a cell
  format-range  Apply formatting to a cell range
  merge         Merge a range of cells
  table         Create a native Excel table over a range
  chart         Create a chart in a worksheet
  pivot         Create a pivot table
  copy-sheet    Copy a worksheet within a workbook
  setup hooks   Install session-start hooks

Run \`excel-axi <command> --help\` for per-command help.`;

export const COMMAND_HELP: Record<string, string> = {
	create: `excel-axi create <file>

Create a new Excel workbook at <file>. The file must not already exist.
The backend (haris-musa/excel-mcp-server) creates an .xlsx with a default Sheet.

Example:
  excel-axi create /data/report.xlsx`,

	"create-sheet": `excel-axi create-sheet <file> <sheet>

Add a new worksheet named <sheet> to an existing workbook at <file>.

Example:
  excel-axi create-sheet /data/report.xlsx Summary`,

	info: `excel-axi info <file>

Show workbook metadata: sheet names, creation info.

Example:
  excel-axi info /data/report.xlsx`,

	read: `excel-axi read <file> <sheet> [start-cell] [end-cell]

Read data from <sheet>. Optional start/end cells in A1 notation.
Returns structured cell data including values and address.

Examples:
  excel-axi read /data/report.xlsx Sheet1
  excel-axi read /data/report.xlsx Sheet1 A1 D20`,

	write: `excel-axi write <file> <sheet> <data-json> [start-cell]

Write a 2D JSON array (array of row arrays) to <sheet>.
Optional <start-cell> defaults to A1.

<data-json> must be a JSON array of arrays.

Examples:
  excel-axi write /data/report.xlsx Sheet1 '[["Name","Score"],["Alice",90]]'
  excel-axi write /data/report.xlsx Sheet1 '[[1,2],[3,4]]' B2`,

	formula: `excel-axi formula <file> <sheet> <cell> <formula>

Apply an Excel formula to a single cell.
The formula must include the leading = sign.

Example:
  excel-axi formula /data/report.xlsx Sheet1 C2 '=SUM(B2:B10)'`,

	"format-range": `excel-axi format-range <file> <sheet> <start-cell> <format-json> [end-cell]

Apply formatting to a cell range. <format-json> is a flat JSON object with
style properties. Supported keys:
  bold (bool), italic (bool), underline (bool)
  font_size (int), font_color (str), bg_color (str)
  border_style (str), border_color (str)
  number_format (str), alignment (str)
  wrap_text (bool)

Example:
  excel-axi format-range /data/report.xlsx Sheet1 A1 '{"bold":true,"font_size":14}' B1`,

	merge: `excel-axi merge <file> <sheet> <start-cell> <end-cell>

Merge cells from <start-cell> to <end-cell> in <sheet>.

Example:
  excel-axi merge /data/report.xlsx Sheet1 A1 C1`,

	table: `excel-axi table <file> <sheet> <range> [name]

Create a native Excel table over <range> (e.g. A1:D5). Optional table [name].

Example:
  excel-axi table /data/report.xlsx Sheet1 A1:B10 ScoreTable`,

	chart: `excel-axi chart <file> <sheet> <data-range> <chart-type> <target-cell>

Create a chart in <sheet>. <chart-type> is one of: line, bar, pie, scatter, area.
<target-cell> is where the top-left corner of the chart is placed.

Example:
  excel-axi chart /data/report.xlsx Sheet1 A1:B4 bar E1`,

	pivot: `excel-axi pivot <file> <sheet> <data-range> <rows-json> <values-json>

Create a pivot table. <rows-json> is a JSON array of field names for row labels.
<values-json> is a JSON array of field names to aggregate.

Example:
  excel-axi pivot /data/report.xlsx Sheet1 A1:C10 '["Name"]' '["Score"]'`,

	"copy-sheet": `excel-axi copy-sheet <file> <src> <dst>

Copy sheet <src> to a new sheet named <dst>.

Example:
  excel-axi copy-sheet /data/report.xlsx Sheet1 Sheet1Copy`,
};
