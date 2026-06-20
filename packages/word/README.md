# @axi-office/word

AXI CLI for Word — create, convert, read, and patch `.docx` files using
[docx](https://www.npmjs.com/package/docx) and
[mammoth](https://www.npmjs.com/package/mammoth). No MCP server or Office install required.

## Disclaimer

Unofficial project — not affiliated with or endorsed by Microsoft Corporation.
All product names and trademarks are property of their respective owners.

## Install

```bash
npm install -g @axi-office/word
```

## Security / path handling

By default, file paths are not sandboxed — any path the calling agent supplies is
resolved relative to the current working directory. When processing untrusted input
(agent-generated specs, user-supplied filenames, or external Markdown), pass
`--base-dir <dir>` to constrain all file operations to a directory:

```bash
word-axi create --base-dir /safe/output out.docx '{"sections":[...]}'
word-axi from-markdown --base-dir /safe notes.md notes.docx
word-axi read --base-dir /docs doc.docx
word-axi info --base-dir /docs doc.docx
word-axi patch --base-dir /safe template.docx '{"name":"Ada"}' --out filled.docx
```

Paths that resolve outside the base directory are rejected with a `SECURITY_ERROR`
before any file system operation.

## Usage

```bash
word-axi <command> [args] [flags]
```

| Command | Description |
|---------|-------------|
| `create <out.docx> <spec-json\|->` | Build a `.docx` from a JSON spec (or stdin) |
| `from-markdown <in.md> <out.docx>` | Convert Markdown to `.docx` |
| `read <in.docx> [--format raw\|html]` | Extract text (raw) or HTML |
| `info <in.docx>` | Report word/paragraph/heading counts |
| `patch <in.docx> <data-json> [--out FILE]` | Replace `{{key}}` placeholders |
| `setup hooks` | Install AXI session-start hooks |

### Create spec format

```json
{
  "title": "My Doc",
  "sections": [
    { "type": "heading", "level": 1, "text": "Introduction" },
    { "type": "paragraph", "text": "Hello world" },
    { "type": "list", "items": ["Item 1", "Item 2"] },
    { "type": "table", "rows": [["H1", "H2"], ["V1", "V2"]] }
  ]
}
```

### Examples

```bash
word-axi create out.docx '{"sections":[{"type":"paragraph","text":"Hi"}]}'
cat spec.json | word-axi create out.docx -
word-axi from-markdown notes.md notes.docx
word-axi read notes.docx --format raw
word-axi info notes.docx
word-axi patch template.docx '{"name":"Ada"}' --out filled.docx
```

### Markdown support

`from-markdown` covers a common subset: `#`/`##`/`###`… headings, paragraphs,
`-`/`*` bullet lists, and `**bold**` / `*italic*` inline spans.

Output is rendered as [TOON](https://github.com/toon-format/toon).

## License

MIT
