# AXI Office

**AXI (Agent eXperience Interface) CLI wrappers for Microsoft Office MCP/CLI backends** — token-efficient, agent-ergonomic.

Provides streamlined interfaces to Excel, PowerPoint, Word, and Outlook for agentic workflows.

## Disclaimer

Unofficial project — not affiliated with or endorsed by Microsoft Corporation.
This package wraps the following open-source MCP servers:
- [haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) (for Excel)
- [GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) (for PowerPoint)
- [softeria/ms-365-mcp-server](https://www.npmjs.com/package/@softeria/ms-365-mcp-server) (for Outlook)

All product names and trademarks are property of their respective owners.

## Status

Early development / scaffolding. API and interfaces subject to change.

---

## Excel (`excel-axi`)

**Backend:** [haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) — Python/uvx, requires `uv` on PATH.

**Install `uv`:** `curl -LsSf https://astral.sh/uv/install.sh | sh`

**Commands:**
- `create <file>` — Create a new Excel workbook
- `create-sheet <file> <sheet>` — Add a worksheet
- `info <file>` — Show workbook metadata
- `read <file> <sheet> [start] [end]` — Read cell data
- `write <file> <sheet> <data-json> [start-cell]` — Write 2D data
- `formula <file> <sheet> <cell> <formula>` — Apply a formula
- `format-range <file> <sheet> <start> <format-json> [--end C]` — Format cells
- `merge <file> <sheet> <start> <end>` — Merge cells
- `table <file> <sheet> <range> [name]` — Create a native table
- `chart <file> <sheet> <range> <type> <target>` — Create a chart
- `pivot <file> <sheet> <range> <rows-json> <values-json> [--agg] [--columns]` — Pivot table
- `copy-sheet <file> <src> <dst>` — Copy a worksheet

---

## PowerPoint (`ppt-axi`)

**Backend:** [GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) — Python/uvx, uses python-pptx.

**Cross-platform:** No Microsoft Office required — python-pptx handles `.pptx` files natively.

**Install `uv`:** `curl -LsSf https://astral.sh/uv/install.sh | sh`

**Note:** The backend is stateful (session-based). Each `ppt-axi` invocation opens, operates, and saves in a single subprocess. State does not persist between invocations beyond the saved file.

**Commands:**
- `create <file>` — Create a new blank presentation
- `info <file>` — Show presentation metadata (slide count, properties)
- `set-props <file> [--title --author --subject --keywords --comments]` — Set document properties
- `from-template <file> <template>` — Create from a `.pptx` template
- `add-slide <file> [--layout N] [--title T] [--color-scheme S]` — Add a slide
- `slide-info <file> <slide-index>` — Show shapes and placeholders on a slide
- `read <file>` — Extract all text from a presentation
- `read-slide <file> <slide-index>` — Extract text from one slide
- `set-placeholder <file> <slide-index> <placeholder-idx> <text>` — Set placeholder text
- `bullets <file> <slide-index> <placeholder-idx> <points-json>` — Add bullet points
- `add-text <file> <slide-index> <text> [position/font flags]` — Add a text box
- `add-image <file> <slide-index> <image-path> [position flags]` — Embed an image
- `add-table <file> <slide-index> <rows> <cols> [--data <json-2d>] [--header-row]` — Add a table
- `add-shape <file> <slide-index> <shape-type> [--fill-color --line-color --text]` — Add a shape
- `add-chart <file> <slide-index> <type> <categories-json> <series-names-json> <series-values-json>` — Add a chart
- `templates` — List available slide templates and color schemes
- `auto-generate <file> <topic> [--slides --type --color-scheme --no-charts --images]` — Auto-generate a deck

---

## Word / Outlook

Early scaffolding — see individual package READMEs.
