---
"@axi-office/excel": major
---

Swap Excel backend to haris-musa/excel-mcp-server; add create-workbook, create-sheet, formula, merge, chart, pivot; requires uv

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
