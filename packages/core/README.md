# @axi-office/core

Shared plumbing for `@axi-office` CLI tools — TOON schema helpers, MCP STDIO client, flag parser, and session hook registration.

## What is AXI?

AXI (Agent eXperience Interface) is a CLI design pattern optimised for use by AI agents. Tools produce structured [TOON](https://github.com/toon-format/toon) output — token-efficient, parseable by LLMs, human-readable in the terminal.

## Installation

```bash
npm install @axi-office/core
```

## API

### `runAxiCli` / `AxiError` (re-exported from `axi-sdk-js`)

```ts
import { runAxiCli, AxiError } from "@axi-office/core";
```

### TOON field-schema helpers

```ts
import { field, pluck, relativeTime, boolYesNo, mapEnum, renderList } from "@axi-office/core";

const schema = [
  field("id"),
  field("title"),
  relativeTime("createdAt", "age"),
  boolYesNo("active"),
];

const output = renderList("worksheets", items, schema);
```

### `parseFlags(args, booleans?)`

```ts
import { parseFlags } from "@axi-office/core";

const { positionals, flags } = parseFlags(process.argv.slice(2), ["verbose"]);
```

### `McpStdioClient`

```ts
import { McpStdioClient } from "@axi-office/core";

const client = new McpStdioClient({
  command: "npx",
  args: ["-y", "@negokaz/mcp-server-excel"],
  name: "excel-mcp",
});

const result = await client.callTool("list_sheets", { filePath: "report.xlsx" });
await client.close();
```

### `setupHooksCommand(marker, binaryNames)`

```ts
import { setupHooksCommand } from "@axi-office/core";

// In your CLI's "setup hooks" subcommand:
return setupHooksCommand("excel-axi", ["excel-axi"]);
```

## License

MIT
