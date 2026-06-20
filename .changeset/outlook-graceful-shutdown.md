---
"@axi-office/outlook": patch
---

Await graceful MCP subprocess shutdown before process exit

The `outlook-axi` bin now calls `await closeAllLiveClients()` from `@axi-office/core` before `process.exit()`. This lets the MCP server flush stderr and exit cleanly, preventing subprocess traceback noise after every command.
