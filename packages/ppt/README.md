# @axi-office/ppt

AXI CLI for PowerPoint — a thin, token-efficient wrapper around
[GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) (MIT).

## Disclaimer

Unofficial project — not affiliated with or endorsed by Microsoft Corporation.
All product names and trademarks are property of their respective owners.

## Prerequisites

The backend is a Python package fetched automatically via `uvx` (from the `uv` toolchain).
Install `uv` once per machine:

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

(other install methods: https://docs.astral.sh/uv/getting-started/installation/ )

On first run, `uvx` will auto-fetch the pinned server
(`office-powerpoint-mcp-server==2.0.7`) from PyPI. Subsequent runs use the local cache.
The backend uses [python-pptx](https://python-pptx.readthedocs.io/) — it is cross-platform
and **no Microsoft PowerPoint installation is required**.

### Backend tradeoff

`Office-PowerPoint-MCP-Server` (python-pptx) was chosen for portability: it runs on Linux,
macOS, and Windows without Microsoft Office. The tradeoff is fidelity — python-pptx cannot
produce animations, SmartArt, or live-rendered output. For those, a COM-automation backend
(Windows-only) would be required.

## Install

```bash
npm install -g @axi-office/ppt
```

## Stateful backend

The backend holds presentations in memory keyed by a `presentation_id` and only writes to
disk on save. To keep the CLI stateless, each `ppt-axi` invocation runs a self-contained
open/create → mutate → save sequence in a single subprocess. State does not persist between
invocations beyond the saved `.pptx` file.

## Usage

```bash
ppt-axi <command> [args] [flags]
```

| Command | Description |
|---------|-------------|
| `create <file>` | Create a new blank presentation |
| `info <file>` | Show presentation metadata (slide count, properties) |
| `set-props <file> [--title T] [--author A] [--subject S] [--keywords K] [--comments C]` | Set core document properties |
| `from-template <file> <template-path>` | Create a presentation from a `.pptx` template |
| `add-slide <file> [--layout N] [--title T] [--color-scheme S]` | Add a slide (layout defaults to 1) |
| `slide-info <file> <slide-index>` | Show shapes and placeholders on a slide (0-based) |
| `read <file>` | Extract all text from a presentation |
| `read-slide <file> <slide-index>` | Extract text from one slide (0-based) |
| `set-placeholder <file> <slide-index> <placeholder-idx> <text>` | Set placeholder text |
| `bullets <file> <slide-index> <placeholder-idx> <points-json>` | Add bullet points to a placeholder |
| `add-text <file> <slide-index> <text> [position/font flags]` | Add a text box |
| `add-image <file> <slide-index> <image-path> [position flags]` | Embed a local image |
| `add-table <file> <slide-index> <rows> <cols> [--data <json-2d>] [--header-row]` | Add a table |
| `add-shape <file> <slide-index> <shape-type> [--text T] [--fill-color R,G,B] [--line-color R,G,B]` | Add a shape |
| `add-chart <file> <slide-index> <chart-type> <categories-json> <series-names-json> <series-values-json>` | Add a chart |
| `templates` | List available slide templates and color schemes |
| `auto-generate <file> <topic> [--slides N] [--type T] [--color-scheme S] [--no-charts] [--images]` | Auto-generate a deck |
| `setup hooks` | Install AXI session-start hooks |

Slide indices are 0-based. Position and size flags (`--left`, `--top`, `--width`,
`--height`) are in inches. Colors are `R,G,B` triples (e.g. `0,112,192`).

### Examples

```bash
# Create a blank presentation
ppt-axi create /data/deck.pptx

# Set document properties
ppt-axi set-props /data/deck.pptx --title "Q4 Review" --author "Alice"

# Add a slide with a title and color scheme
ppt-axi add-slide /data/deck.pptx --layout 2 --title "Agenda" --color-scheme modern_blue

# Inspect a slide's shapes and placeholders
ppt-axi slide-info /data/deck.pptx 0

# Set placeholder text and add bullets
ppt-axi set-placeholder /data/deck.pptx 0 0 "Welcome to Q4 Review"
ppt-axi bullets /data/deck.pptx 1 1 '["Revenue up 20%","Costs reduced","New markets"]'

# Add a text box (positions in inches)
ppt-axi add-text /data/deck.pptx 0 "Bold Title" --left 2 --top 1 --bold --font-size 32

# Embed an image
ppt-axi add-image /data/deck.pptx 2 /assets/logo.png --left 0.5 --top 0.5 --width 2

# Add a table with header row
ppt-axi add-table /data/deck.pptx 1 2 3 --data '[["Name","Score","Grade"],["Alice","90","A"]]' --header-row

# Add a shape
ppt-axi add-shape /data/deck.pptx 0 rectangle --text "Key Point" --fill-color 0,112,192

# Add a chart (types: column, bar, line, pie)
ppt-axi add-chart /data/deck.pptx 2 column '["Q1","Q2","Q3"]' '["Revenue"]' '[[100,200,300]]'

# Read text back out
ppt-axi read /data/deck.pptx
ppt-axi read-slide /data/deck.pptx 2

# List templates / color schemes
ppt-axi templates

# Auto-generate a full deck from a topic
ppt-axi auto-generate /data/deck.pptx "AI trends in 2025" --slides 8 --type academic
```

`auto-generate` flags: `--type` is one of `business`, `academic`, `creative` (default
`business`); `--color-scheme` is one of `modern_blue`, `corporate_gray`, `elegant_green`,
`warm_red` (default `modern_blue`); `--slides` defaults to `5`; `--no-charts` excludes
charts; `--images` includes images.

File paths should be absolute. Output is rendered as
[TOON](https://github.com/toon-format/toon) for token-efficient agent consumption.

**Security note:** Paths are not sandboxed — the server reads/writes any path your user
account can access; pass absolute paths.

## Pinned server version

This package pins `office-powerpoint-mcp-server==2.0.7`. Audit the upstream
[repository](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server)
before bumping the pin in `src/client.ts`.

## License

MIT
