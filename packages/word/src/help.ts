/**
 * Why: runAxiCli needs a top-level help string and per-command help lookup; centralizing
 * them keeps usage text out of the command modules.
 * What: Exports TOP_LEVEL_HELP and a COMMAND_HELP record keyed by command name.
 * Test: Assert TOP_LEVEL_HELP contains "word-axi" and COMMAND_HELP.create is non-empty.
 */

export const TOP_LEVEL_HELP = `word-axi — AXI CLI for Word (.docx)

Usage: word-axi <command> [args] [flags]

Commands:
  create         Build a .docx from a JSON spec
  from-markdown  Build a .docx from a Markdown file
  read           Extract text or HTML from a .docx
  info           Report word/paragraph/heading counts
  patch          Replace {{key}} placeholders in a .docx
  setup hooks    Install session-start hooks

Run \`word-axi <command> --help\` for per-command help.`;

export const COMMAND_HELP: Record<string, string> = {
  create: `word-axi create <out.docx> <spec-json|->

Build a .docx from a JSON spec. Pass "-" to read the spec from stdin.
Spec: { "title"?: string, "sections": Section[] }
Section types:
  { "type": "heading", "level": 1-6, "text": "..." }
  { "type": "paragraph", "text": "..." }
  { "type": "list", "items": ["..."] }
  { "type": "table", "rows": [["H1","H2"],["v1","v2"]] }`,
  "from-markdown": `word-axi from-markdown <in.md> <out.docx>

Convert a Markdown file to .docx. Supports # / ## / ### headings,
paragraphs, "- "/"* " bullet lists, **bold** and *italic* inline spans.`,
  read: `word-axi read <in.docx> [--format raw|html]

Extract document content. --format raw (default) returns plain text;
--format html returns HTML produced by mammoth.`,
  info: `word-axi info <in.docx>

Report counts: words, paragraphs, and headings (best-effort from raw text).`,
  patch: `word-axi patch <in.docx> <data-json> [--out FILE]

Replace {{key}} placeholders with values from <data-json>.
Without --out the input file is overwritten in place.`,
};
