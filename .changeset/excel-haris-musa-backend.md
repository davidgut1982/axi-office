---
"@axi-office/excel": major
---

Swap Excel backend to haris-musa/excel-mcp-server; add create-workbook, create-sheet, formula, merge, chart, pivot; requires uv

**Breaking changes:**
- Backend changed from `@negokaz/excel-mcp-server` (Node/npx) to `haris-musa/excel-mcp-server` (Python/uvx). Requires `uv` on PATH (`curl -LsSf https://astral.sh/uv/install.sh | sh`).
- `sheets <file>` command removed; replaced by `info <file>` (calls `get_workbook_metadata`).
- `write <file> <sheet> <range> <values-json>` changed to `write <file> <sheet> <data-json> [start-cell]` — the `<range>` positional is gone; data is a 2D JSON array of rows.
- `format-range` now takes a flat style object (`{"bold":true,"font_size":14}`) instead of the negokaz 2D per-cell styles array.
- `create-table` CLI command renamed from `create-table` to `table`.
- `copy-sheet` now calls `copy_worksheet` (different MCP tool name and arg keys).

**New commands:**
- `create <file>` — create a workbook from scratch (fills the primary gap)
- `create-sheet <file> <sheet>` — add a worksheet
- `formula <file> <sheet> <cell> <formula>` — apply an Excel formula
- `merge <file> <sheet> <start> <end>` — merge cells
- `chart <file> <sheet> <data-range> <type> <target-cell>` — create a chart
- `pivot <file> <sheet> <data-range> <rows-json> <values-json> [--agg]` — create a pivot table
