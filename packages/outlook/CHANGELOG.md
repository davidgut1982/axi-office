# @axi-office/outlook

## 0.2.1

### Patch Changes

- 192dc1c: Await graceful MCP subprocess shutdown before process exit

  The `outlook-axi` bin now calls `await closeAllLiveClients()` from `@axi-office/core` before `process.exit()`. This lets the MCP server flush stderr and exit cleanly, preventing subprocess traceback noise after every command.

- Updated dependencies [192dc1c]
  - @axi-office/core@0.2.1

## 0.2.0

### Minor Changes

- 3bcc417: Add `@axi-office/outlook` — AXI CLI wrapping softeria/ms-365-mcp-server (`--preset outlook`).

  Commands: `mail-list`, `mail-get`, `mail-send`, `folders`, `cal-list`, `cal-view`,
  `cal-create`, `contacts-list`, and `login`, plus `setup hooks`. Renders results as TOON.

### Patch Changes

- be99b3b: Outlook calendar: normalise bare dates to full ISO 8601 datetimes in cal-view.

  - calViewCommand now accepts either full ISO 8601 datetimes (passed through unchanged) or bare YYYY-MM-DD dates (expanded: start → T00:00:00Z, end → T23:59:59Z).
  - Invalid strings (neither a bare date nor an ISO datetime) throw an AxiError with a clear format hint.
  - Updated help text and added tests for bare-date expansion, mixed input, and invalid input.

- 5c17aef: Security remediations for public release: pin upstream MCP server versions (supply-chain TOFU fix), add opt-in path sandbox to word commands (--base-dir), harden outlook README with Entra app registration and token security guidance, mark callHttpTool @internal with SSRF note, add disclaimer to all READMEs.
- Updated dependencies [5d778e0]
- Updated dependencies [641cad8]
- Updated dependencies [5c17aef]
  - @axi-office/core@0.2.0
