---
"@axi-office/word": patch
---

Word package: add round-trip patch test, document PatchType.PARAGRAPH behavior, rename headings heuristic field.

- patch: add round-trip test that exercises placeholder-only paragraphs with surrounding static text in separate paragraphs, locking in and documenting that PatchType.PARAGRAPH replaces the entire containing paragraph (so the placeholder must be the sole paragraph content). Update help text with this constraint.
- info: rename the `headings` output field to `headings_estimated` with an inline comment clarifying it is a heuristic. The old field name was misleading; the count may over- or under-count.
- word-axi bin: add `process.exit(process.exitCode ?? 0)` after `runAxiCli` to prevent the process from hanging due to open MCP stdio handles.
