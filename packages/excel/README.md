# @axi-office/excel

AXI CLI for Excel — a thin, token-efficient wrapper around
[haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) (MIT).

## Disclaimer

Unofficial project — not affiliated with or endorsed by Microsoft Corporation.
All product names and trademarks are property of their respective owners.

## Prerequisites

The backend is a Python package fetched automatically via `uvx` (from the `uv` toolchain).
Install `uv` once per machine:

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

On first run, `uvx` will auto-fetch the pinned server
(`excel-mcp-server@0.1.8`) from PyPI. Subsequent runs use the local cache.
No Microsoft Excel installation is required.

## Install

```bash
npm install -g @axi-office/excel
```

## Usage

```bash
excel-axi <command> [args] [flags]
```

| Command | Description |
|---------|-------------|
| `create <file>` | Create a new Excel workbook from scratch |
| `create-sheet <file> <sheet>` | Add a new worksheet to a workbook |
| `info <file>` | Show workbook metadata (sheet names, etc.) |
| `read <file> <sheet> [start] [end]` | Read data from a worksheet |
| `write <file> <sheet> <data-json> [start-cell]` | Write a 2D JSON array to a worksheet |
| `formula <file> <sheet> <cell> <formula>` | Apply an Excel formula to a cell |
| `format-range <file> <sheet> <start> <format-json> [end]` | Apply formatting to a range |
| `merge <file> <sheet> <start> <end>` | Merge a range of cells |
| `table <file> <sheet> <range> [name]` | Create a native Excel table |
| `chart <file> <sheet> <data-range> <type> <target-cell>` | Create a chart |
| `pivot <file> <sheet> <data-range> <rows-json> <values-json>` | Create a pivot table |
| `copy-sheet <file> <src> <dst>` | Copy a worksheet within a workbook |
| `setup hooks` | Install AXI session-start hooks |

### Examples

```bash
# Create a workbook from scratch
excel-axi create /data/report.xlsx

# Add a sheet
excel-axi create-sheet /data/report.xlsx Summary

# Show metadata
excel-axi info /data/report.xlsx

# Write data (2D array of rows)
excel-axi write /data/report.xlsx Sheet1 '[["Name","Score"],["Alice",99],["Bob",85]]'

# Read data
excel-axi read /data/report.xlsx Sheet1
excel-axi read /data/report.xlsx Sheet1 A1 D20

# Apply a formula
excel-axi formula /data/report.xlsx Sheet1 C2 '=SUM(B2:B4)'

# Format cells (flat style object)
excel-axi format-range /data/report.xlsx Sheet1 A1 '{"bold":true,"font_size":14}' C1

# Merge cells
excel-axi merge /data/report.xlsx Sheet1 A1 C1

# Create a table
excel-axi table /data/report.xlsx Sheet1 A1:C4 ScoreTable

# Create a chart (types: line, bar, pie, scatter, area)
excel-axi chart /data/report.xlsx Sheet1 A1:B4 bar E1

# Create a pivot table
excel-axi pivot /data/report.xlsx Sheet1 A1:C10 '["Name"]' '["Score"]'

# Copy a sheet
excel-axi copy-sheet /data/report.xlsx Sheet1 Sheet1Copy
```

File paths should be absolute. Output is rendered as
[TOON](https://github.com/toon-format/toon) for token-efficient agent consumption.

## Pinned server version

This package pins `excel-mcp-server@0.1.8`. Audit the upstream
[CHANGELOG](https://github.com/haris-musa/excel-mcp-server/blob/main/CHANGELOG.md)
before bumping the pin in `src/client.ts`.

## License

MIT
