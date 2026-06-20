# @axi-office/excel

AXI CLI for Excel — a thin, token-efficient wrapper around
[negokaz/excel-mcp-server](https://www.npmjs.com/package/@negokaz/excel-mcp-server).

## Disclaimer

Unofficial project — not affiliated with or endorsed by Microsoft Corporation.
This package wraps the following open-source MCP servers:
- [negokaz/excel-mcp-server](https://www.npmjs.com/package/@negokaz/excel-mcp-server) (for Excel)

All product names and trademarks are property of their respective owners.

## Install

```bash
npm install -g @axi-office/excel
```

The CLI auto-fetches the pinned server (`@negokaz/excel-mcp-server@0.12.0`) via npx on first run (network required once). You may pre-install a pinned global to avoid auto-fetch: `npm install -g @negokaz/excel-mcp-server@0.12.0`

## Usage

```bash
excel-axi <command> [args] [flags]
```

| Command | Description |
|---------|-------------|
| `sheets <file>` | Describe all sheets in a workbook |
| `read <file> <sheet> [range] [--limit N]` | Read cell values from a sheet |
| `write <file> <sheet> <range> <values-json>` | Write a 2D JSON array to a range |
| `create-table <file> <sheet> <range> [name]` | Create a table over a range |
| `copy-sheet <file> <src> <dst>` | Copy a sheet within a workbook |
| `format-range <file> <sheet> <range> <format-json>` | Apply formatting to a range |
| `setup hooks` | Install AXI session-start hooks |

### Examples

```bash
excel-axi sheets /data/report.xlsx
excel-axi read /data/report.xlsx Sheet1 A1:D20 --limit 50
excel-axi write /data/report.xlsx Sheet1 A1:B2 '[["Name","Score"],["Ada",99]]'
excel-axi create-table /data/report.xlsx Sheet1 A1:B2 Results
excel-axi copy-sheet /data/report.xlsx Sheet1 Sheet1Copy
excel-axi format-range /data/report.xlsx Sheet1 A1:B1 '{"bold":true}'
```

File paths must be absolute. Output is rendered as
[TOON](https://github.com/toon-format/toon) for token-efficient agent consumption.

## License

MIT
