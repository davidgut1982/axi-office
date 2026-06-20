# @axi-office/word

## 0.2.1

### Patch Changes

- Updated dependencies [192dc1c]
  - @axi-office/core@0.2.1

## 0.2.0

### Minor Changes

- 3bcc417: Add `@axi-office/word` — AXI CLI for `.docx` files using docx and mammoth (no MCP server).

  Commands: `create` (JSON spec or stdin), `from-markdown`, `read` (raw/html), `info`, and
  `patch` ({{key}} placeholder substitution), plus `setup hooks`.

### Patch Changes

- be99b3b: Word package: add round-trip patch test, document PatchType.PARAGRAPH behavior, rename headings heuristic field.

  - patch: add round-trip test that exercises placeholder-only paragraphs with surrounding static text in separate paragraphs, locking in and documenting that PatchType.PARAGRAPH replaces the entire containing paragraph (so the placeholder must be the sole paragraph content). Update help text with this constraint.
  - info: rename the `headings` output field to `headings_estimated` with an inline comment clarifying it is a heuristic. The old field name was misleading; the count may over- or under-count.
  - word-axi bin: add `process.exit(process.exitCode ?? 0)` after `runAxiCli` to prevent the process from hanging due to open MCP stdio handles.

- 5c17aef: Security remediations for public release: pin upstream MCP server versions (supply-chain TOFU fix), add opt-in path sandbox to word commands (--base-dir), harden outlook README with Entra app registration and token security guidance, mark callHttpTool @internal with SSRF note, add disclaimer to all READMEs.
- Updated dependencies [5d778e0]
- Updated dependencies [641cad8]
- Updated dependencies [5c17aef]
  - @axi-office/core@0.2.0
