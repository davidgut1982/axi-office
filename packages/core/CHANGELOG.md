# @axi-office/core

## 0.2.0

### Minor Changes

- 5d778e0: Add `@axi-office/core` — shared plumbing for `@axi-office` CLI tools.

  Provides:

  - TOON field-schema helpers (`field`, `pluck`, `relativeTime`, `boolYesNo`, `mapEnum`, `renderList`, `renderDetail`, `renderHelp`) ported from `gh-axi/dist/src/toon.js` to strict TypeScript.
  - `parseFlags(args, booleans?)` and related arg utilities ported from `lore-axi/src/args.ts`.
  - `McpStdioClient` — lazy-connect MCP STDIO client wrapping `@modelcontextprotocol/sdk` for use by excel/word/outlook packages.
  - `callHttpTool()` — single-round-trip HTTP MCP helper.
  - `setupHooksCommand(marker, binaryNames)` — wraps `installSessionStartHooks` for consistent session hook setup.
  - Re-exports `runAxiCli`, `AxiError`, `installSessionStartHooks` from `axi-sdk-js`.

  Bug fixes (adversarial review):

  - `McpStdioClient`: clear `connectPromise` on connect failure so subsequent calls retry rather than permanently awaiting a rejected promise.
  - `McpStdioClient`: replace `process.once("exit", ...)` with SIGINT/SIGTERM graceful handlers plus a synchronous SIGKILL backstop on `"exit"`; handlers are registered once at module level to avoid accumulation across instances.
  - `parseFlags`: treat `"--"` as end-of-flags sentinel; remaining tokens become positionals and no spurious `flags[""]` key is produced.
  - Remove dead `scripts/lint.mjs`; lint runs via `biome check src/ test/` directly.

### Patch Changes

- 641cad8: Harden McpStdioClient: POSIX exit codes, opt-in auto-exit, close/connect race fix, subprocess observability warn, connect timeout, expanded tests
- 5c17aef: Security remediations for public release: pin upstream MCP server versions (supply-chain TOFU fix), add opt-in path sandbox to word commands (--base-dir), harden outlook README with Entra app registration and token security guidance, mark callHttpTool @internal with SSRF note, add disclaimer to all READMEs.
