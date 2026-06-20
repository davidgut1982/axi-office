---
"@axi-office/core": minor
---

Add `@axi-office/core` — shared plumbing for `@axi-office` CLI tools.

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
