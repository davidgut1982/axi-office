# @axi-office/excel

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
