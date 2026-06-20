# @axi-office/excel

## 1.0.0

### Major Changes

- 192dc1c: Swap Excel backend to haris-musa/excel-mcp-server; add create-workbook, create-sheet, formula, merge, chart, pivot; requires uv

  **Breaking changes:**

  - Backend changed from `@negokaz/excel-mcp-server` (Node/npx) to `haris-musa/excel-mcp-server` (Python/uvx). Requires `uv` on PATH (`curl -LsSf https://astral.sh/uv/install.sh | sh` or https://docs.astral.sh/uv/getting-started/installation/ ).
  - `sheets <file>` command removed; replaced by `info <file>` (calls `get_workbook_metadata`).
  - `write <file> <sheet> <range> <values-json>` changed to `write <file> <sheet> <data-json> [start-cell]` — the `<range>` positional is gone; data is a 2D JSON array of rows.
  - `format-range` now takes a flat style object (`{"bold":true,"font_size":14}`) instead of the negokaz 2D per-cell styles array. The optional end-cell argument moved from the 5th positional to `--end <cell>`.
  - `create-table` CLI command renamed from `create-table` to `table`.
  - `copy-sheet` now calls `copy_worksheet` (different MCP tool name and arg keys).

  **New commands:**

  - `create <file>` — create a workbook from scratch (fills the primary gap)
  - `create-sheet <file> <sheet>` — add a worksheet
  - `formula <file> <sheet> <cell> <formula>` — apply an Excel formula
  - `merge <file> <sheet> <start> <end>` — merge cells
  - `chart <file> <sheet> <data-range> <type> <target-cell>` — create a chart
  - `pivot <file> <sheet> <data-range> <rows-json> <values-json> [--agg] [--columns <json-array>]` — create a pivot table

  **Fixes:**

  - `info`: removed hardcoded `include_ranges: false`; lets the server use its own default.
  - `write`: added 2D-array validation — flat `[val, val]` arrays now produce a clear error instead of silently sending bad data to the server.
  - `format-range`: end-cell moved from 5th positional to `--end <cell>` flag (unambiguous when format-json contains spaces).
  - `pivot`: added optional `--columns <json-array>` flag for column grouping fields.
  - Bin now awaits `closeAllLiveClients()` before `process.exit()`, eliminating the `ValueError: I/O operation on closed file` / "Service stopped." tracebacks that printed to stderr after every command.

### Patch Changes

- Updated dependencies [192dc1c]
  - @axi-office/core@0.2.1

## 0.2.0

### Minor Changes

- 3bcc417: Add `@axi-office/excel` — AXI CLI wrapping negokaz/excel-mcp-server.

  Commands: `sheets`, `read`, `write`, `create-table`, `copy-sheet`, `format-range`,
  plus `setup hooks`. Spawns `@negokaz/excel-mcp-server` on demand via `npx` and renders
  results as TOON.

### Patch Changes

- be99b3b: Fix excel format-range and read commands to use real excel-mcp-server arg schemas.

  - format-range: accept `--styles` as a 2D JSON array of per-cell style objects (matching the range grid rows×cols) instead of a flat object spread. Add range dimension validation that throws a clear AxiError when the array shape does not match. Update help text with a correct 2D example.
  - read: remove bogus `knownPagingRanges` arg (not a real parameter); replace `--limit` with `--formula` (showFormula) and `--style` (showStyle) boolean flags that match the real `excel_read_sheet` schema. Document EXCEL_MCP_PAGING_CELLS_LIMIT env var in help text.

- 5c17aef: Security remediations for public release: pin upstream MCP server versions (supply-chain TOFU fix), add opt-in path sandbox to word commands (--base-dir), harden outlook README with Entra app registration and token security guidance, mark callHttpTool @internal with SSRF note, add disclaimer to all READMEs.
- Updated dependencies [5d778e0]
- Updated dependencies [641cad8]
- Updated dependencies [5c17aef]
  - @axi-office/core@0.2.0
