---
"@axi-office/core": patch
---

Add `closeAllLiveClients()` for graceful MCP subprocess shutdown

Export a new `closeAllLiveClients(): Promise<void>` from `@axi-office/core` that calls `close()` on every client in the live-clients registry (via `Promise.allSettled`, swallowing individual errors). This allows CLI bins to await clean subprocess teardown before calling `process.exit()`, eliminating the `ValueError: I/O operation on closed file` / "Service stopped." tracebacks that printed to stderr after every command when the Python MCP server's stderr pipe was closed by Node before the child had a chance to exit.
