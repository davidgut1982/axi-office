/**
 * Why: runAxiCli needs a top-level help string and per-command help lookup; keeping
 * them in one module avoids scattering usage strings across command files.
 * What: Exports TOP_LEVEL_HELP and a COMMAND_HELP record keyed by command name.
 * Test: Assert TOP_LEVEL_HELP contains "excel-axi" and COMMAND_HELP.sheets is a
 * non-empty string.
 */

export const TOP_LEVEL_HELP = `excel-axi — AXI CLI for Excel (via negokaz/excel-mcp-server)

Usage: excel-axi <command> [args] [flags]

Commands:
  sheets        Describe sheets in a workbook
  read          Read a range from a sheet
  write         Write JSON values to a range
  create-table  Create a table over a range
  copy-sheet    Copy a sheet within a workbook
  format-range  Apply formatting to a range
  setup hooks   Install session-start hooks

Run \`excel-axi <command> --help\` for per-command help.`;

export const COMMAND_HELP: Record<string, string> = {
	sheets: `excel-axi sheets <file>

Describe all sheets in the workbook at <file>.`,
	read: `excel-axi read <file> <sheet> [range] [--formula] [--style]

Read cell values from <sheet>. Optional A1 <range> (e.g. A1:C10).
--formula   show cell formulas instead of computed values
--style     include style information for each cell

Tip: set EXCEL_MCP_PAGING_CELLS_LIMIT=N on the server to control page size (default 2000).`,
	write: `excel-axi write <file> <sheet> <range> <values-json>

Write a 2D array of values to <range>.
<values-json> is a JSON array of rows, e.g. '[["a","b"],[1,2]]'.`,
	"create-table": `excel-axi create-table <file> <sheet> <range> [name]

Create a table over <range>. Optional table [name].`,
	"copy-sheet": `excel-axi copy-sheet <file> <src> <dst>

Copy sheet <src> to a new sheet named <dst>.`,
	"format-range": `excel-axi format-range <file> <sheet> <range> <styles-json>

Apply formatting to <range>. <styles-json> is a 2D JSON array of per-cell style objects
whose dimensions must exactly match the range grid (rows × cols).

Each element is either a style object or null (no change for that cell).

Example — A1:B2 (2 rows × 2 cols):
  excel-axi format-range book.xlsx Sheet1 A1:B2 '[[{"font":{"bold":true}},null],[null,null]]'`,
};
