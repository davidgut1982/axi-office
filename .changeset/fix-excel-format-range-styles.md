---
"@axi-office/excel": patch
---

Fix excel format-range and read commands to use real excel-mcp-server arg schemas.

- format-range: accept `--styles` as a 2D JSON array of per-cell style objects (matching the range grid rows×cols) instead of a flat object spread. Add range dimension validation that throws a clear AxiError when the array shape does not match. Update help text with a correct 2D example.
- read: remove bogus `knownPagingRanges` arg (not a real parameter); replace `--limit` with `--formula` (showFormula) and `--style` (showStyle) boolean flags that match the real `excel_read_sheet` schema. Document EXCEL_MCP_PAGING_CELLS_LIMIT env var in help text.
