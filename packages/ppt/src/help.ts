/**
 * Why: runAxiCli needs a top-level help string and per-command help lookup; keeping
 * them in one module avoids scattering usage strings across command files.
 * What: Exports TOP_LEVEL_HELP and a COMMAND_HELP record keyed by command name.
 * Test: Assert TOP_LEVEL_HELP contains "ppt-axi" and COMMAND_HELP.create is a
 * non-empty string.
 */

export const TOP_LEVEL_HELP = `ppt-axi — AXI CLI for PowerPoint (via GongRzhe/Office-PowerPoint-MCP-Server, run with uvx)

Requires: uv/uvx on PATH. Install: curl -LsSf https://astral.sh/uv/install.sh | sh
Backend uses python-pptx — cross-platform, no Microsoft Office required.

Usage: ppt-axi <command> [args] [flags]

Commands:
  create            Create a new blank presentation
  info              Show presentation metadata (slide count, properties)
  set-props         Set core document properties (title, author, etc.)
  from-template     Create a presentation from a .pptx template
  add-slide         Add a slide (with optional layout, title, color scheme)
  slide-info        Show shapes and content on a specific slide
  read              Extract all text from a presentation
  read-slide        Extract text from a single slide
  set-placeholder   Set text in a layout placeholder
  bullets           Add bullet points to a placeholder
  add-text          Add a text box at a specific position
  add-image         Embed an image at a specific position
  add-table         Add a table to a slide
  add-shape         Add a shape (rectangle, oval, etc.)
  add-chart         Add a chart with data series
  templates         List available slide templates / color schemes
  auto-generate     Auto-generate a complete presentation from a topic

Run \`ppt-axi <command> --help\` for per-command help.`;

export const COMMAND_HELP: Record<string, string> = {
	create: `ppt-axi create <file>

Create a new blank PowerPoint presentation at <file>.
Requires uv/uvx on PATH (curl -LsSf https://astral.sh/uv/install.sh | sh)

Example:
  ppt-axi create /data/deck.pptx`,

	info: `ppt-axi info <file>

Show presentation metadata: slide count, title, author, and core properties.

Example:
  ppt-axi info /data/deck.pptx`,

	"set-props": `ppt-axi set-props <file> [--title T] [--author A] [--subject S] [--keywords K] [--comments C]

Set core document properties on an existing presentation.
At least one property flag is required.

Example:
  ppt-axi set-props /data/deck.pptx --title "Q4 Review" --author "Alice"`,

	"from-template": `ppt-axi from-template <file> <template-path>

Create a new presentation at <file> using <template-path> as the design template.

Example:
  ppt-axi from-template /data/deck.pptx /templates/corporate.pptx`,

	"add-slide": `ppt-axi add-slide <file> [--layout N] [--title T] [--color-scheme S]

Add a new slide to the presentation. Layout index defaults to 1 (title+content).
Use \`ppt-axi templates\` to list available color schemes.

Examples:
  ppt-axi add-slide /data/deck.pptx
  ppt-axi add-slide /data/deck.pptx --layout 2 --title "Agenda" --color-scheme modern_blue`,

	"slide-info": `ppt-axi slide-info <file> <slide-index>

Show all shapes, placeholders, and content on a specific slide.
<slide-index> is 0-based.

Example:
  ppt-axi slide-info /data/deck.pptx 0`,

	read: `ppt-axi read <file>

Extract all text from a presentation with slide-level context.

Example:
  ppt-axi read /data/deck.pptx`,

	"read-slide": `ppt-axi read-slide <file> <slide-index>

Extract text from a single slide. <slide-index> is 0-based.

Example:
  ppt-axi read-slide /data/deck.pptx 2`,

	"set-placeholder": `ppt-axi set-placeholder <file> <slide-index> <placeholder-idx> <text>

Set text in a layout placeholder. Use \`ppt-axi slide-info\` to find placeholder indices.

Example:
  ppt-axi set-placeholder /data/deck.pptx 0 0 "Welcome to Q4 Review"`,

	bullets: `ppt-axi bullets <file> <slide-index> <placeholder-idx> <points-json>

Add bullet points to a placeholder. <points-json> must be a JSON array of strings.

Example:
  ppt-axi bullets /data/deck.pptx 1 1 '["Revenue up 20%","Costs reduced","New markets"]'`,

	"add-text": `ppt-axi add-text <file> <slide-index> <text> [--left=1] [--top=1] [--width=4] [--height=2]
  [--font-size N] [--font-name N] [--bold] [--italic] [--underline]
  [--color R,G,B] [--align left|center|right]

Add a text box at a position (in inches). Defaults: left=1, top=1, width=4, height=2.

Examples:
  ppt-axi add-text /data/deck.pptx 0 "Hello World"
  ppt-axi add-text /data/deck.pptx 0 "Bold Title" --left 2 --top 1 --bold --font-size 32`,

	"add-image": `ppt-axi add-image <file> <slide-index> <image-path> [--left=1] [--top=1] [--width W] [--height H]

Embed a local image file at a position on a slide (coordinates in inches).

Example:
  ppt-axi add-image /data/deck.pptx 2 /assets/logo.png --left 0.5 --top 0.5 --width 2`,

	"add-table": `ppt-axi add-table <file> <slide-index> <rows> <cols>
  [--left=1] [--top=1] [--width=8] [--height=4]
  [--data <json-2d>] [--header-row]

Add a table with <rows> rows and <cols> columns. Position and size in inches.
Optional --data provides initial cell content as a 2D JSON string array.
Optional --header-row applies header formatting to the first row.

Examples:
  ppt-axi add-table /data/deck.pptx 1 4 3
  ppt-axi add-table /data/deck.pptx 1 2 3 --data '[["Name","Score","Grade"],["Alice","90","A"]]' --header-row`,

	"add-shape": `ppt-axi add-shape <file> <slide-index> <shape-type>
  [--left=1] [--top=1] [--width=2] [--height=2]
  [--text T] [--fill-color R,G,B] [--line-color R,G,B]

Add a shape to a slide. Position and size in inches.

Example:
  ppt-axi add-shape /data/deck.pptx 0 rectangle --text "Key Point" --fill-color 0,112,192`,

	"add-chart": `ppt-axi add-chart <file> <slide-index> <chart-type> <categories-json> <series-names-json> <series-values-json>
  [--left=1] [--top=1] [--width=8] [--height=5] [--title T] [--legend-position=right]

Add a chart. chart-type: column, bar, line, pie.
All JSON args must be valid JSON arrays.

Example:
  ppt-axi add-chart /data/deck.pptx 2 column '["Q1","Q2","Q3"]' '["Revenue"]' '[[100,200,300]]'`,

	templates: `ppt-axi templates

List available slide templates and color schemes supported by the backend.

Example:
  ppt-axi templates`,

	"auto-generate": `ppt-axi auto-generate <file> <topic> [--slides=5] [--type=business] [--color-scheme=modern_blue]
  [--no-charts] [--images]

Auto-generate a complete presentation from a topic description.
  --type:           business, academic, creative (default: business)
  --color-scheme:   modern_blue, corporate_gray, elegant_green, warm_red (default: modern_blue)
  --slides:         number of slides (default: 5)
  --no-charts:      exclude charts from generated content
  --images:         include images in generated content

Example:
  ppt-axi auto-generate /data/deck.pptx "AI trends in 2025" --slides 8 --type academic`,
};
