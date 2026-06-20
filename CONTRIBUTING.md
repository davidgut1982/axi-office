# Contributing to axi-office

This is a pnpm monorepo of AXI CLI tools for Office suite integrations. Each package under `packages/` follows the same structure and publishes to npm as `@axi-office/<name>`.

## Adding a New `@axi-office/<service>` Package

The steps below are mechanical — follow them and your new CLI will slot into the monorepo build, test, lint, and release pipeline automatically.

### 1. Scaffold the directory

```bash
mkdir -p packages/<service>/src/commands
mkdir -p packages/<service>/src/bin
mkdir -p packages/<service>/test
```

### 2. Create `packages/<service>/package.json`

```json
{
  "name": "@axi-office/<service>",
  "version": "0.1.0",
  "description": "AXI CLI for <service>",
  "keywords": ["axi", "agent", "cli", "<service>"],
  "license": "MIT",
  "type": "module",
  "engines": { "node": ">=20" },
  "publishConfig": { "access": "public" },
  "files": ["dist", "README.md"],
  "bin": { "<service>-axi": "./dist/bin/<service>-axi.js" },
  "exports": {
    ".": {
      "import": "./dist/index.js",
      "types": "./dist/index.d.ts"
    }
  },
  "scripts": {
    "build": "tsc -p tsconfig.json",
    "test": "vitest run",
    "test:watch": "vitest",
    "typecheck": "tsc --noEmit",
    "lint": "biome check .",
    "lint:fix": "biome check --write ."
  },
  "dependencies": {
    "@axi-office/core": "workspace:*",
    "axi-sdk-js": "^0.1.7"
  },
  "devDependencies": {
    "@biomejs/biome": "^1.9.0",
    "@types/node": "^22.0.0",
    "typescript": "^5.7.0",
    "vitest": "^3.0.0"
  }
}
```

### 3. Create `packages/<service>/tsconfig.json`

```json
{
  "extends": "../../tsconfig.json",
  "compilerOptions": {
    "outDir": "dist",
    "rootDir": "src"
  },
  "include": ["src/**/*.ts"],
  "exclude": ["node_modules", "dist", "test"]
}
```

### 4. Create the bin entry point

`packages/<service>/src/bin/<service>-axi.ts`:

```typescript
#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { runAxiCli, AxiError, setupHooksCommand } from "@axi-office/core";

function readVersion(): string {
  try {
    const here = dirname(fileURLToPath(import.meta.url));
    const pkg = JSON.parse(readFileSync(join(here, "../../package.json"), "utf8"));
    return typeof pkg.version === "string" ? pkg.version : "0.0.0";
  } catch {
    return "0.0.0";
  }
}

await runAxiCli({
  description: "AXI CLI for <service>",
  version: readVersion(),
  topLevelHelp: `<service>-axi — AXI CLI for <service>\n\nRun \`<service>-axi <command> --help\` for per-command help.\n`,
  getCommandHelp: (_command: string) => null,
  home: (_args: string[]) => ({
    bin: "<service>-axi",
    description: "AXI CLI for <service>",
    version: readVersion(),
  }),
  commands: {
    // TODO: add your commands here

    setup: async (args: string[]) => {
      const sub = args[0];
      if (sub === "hooks") {
        return setupHooksCommand("<service>-axi", ["<service>-axi"]);
      }
      throw new AxiError("setup subcommand is required", "VALIDATION_ERROR", [
        "<service>-axi setup hooks    — Install session-start hooks",
      ]);
    },
  },
});
```

### 5. Create `packages/<service>/vitest.config.ts`

```typescript
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["test/**/*.test.ts"],
    environment: "node",
  },
});
```

### 6. Wire the MCP STDIO client

If your service talks to an MCP server subprocess:

```typescript
// src/client.ts
import { McpStdioClient } from "@axi-office/core";

let _client: McpStdioClient | undefined;

export function getClient(): McpStdioClient {
  if (!_client) {
    _client = new McpStdioClient({
      command: "npx",
      args: ["-y", "@vendor/mcp-server-<service>"],
      name: "<service>-mcp",
    });
  }
  return _client;
}
```

### 7. Install and build

```bash
pnpm install
pnpm --filter @axi-office/<service> build
pnpm --filter @axi-office/<service> test
```

### 8. Add a changeset

```bash
pnpm changeset
# Select @axi-office/<service>, choose patch/minor/major, write a summary
```

### 9. Open a PR

The CI workflow runs `pnpm -r build`, `pnpm -r test`, and `pnpm -r lint` on every PR. Once merged to `main`, the Release workflow will create a "Release PR" or publish automatically.

## Monorepo Commands

| Command | What it does |
|---------|-------------|
| `pnpm -r build` | Build all packages |
| `pnpm -r test` | Run all tests |
| `pnpm -r lint` | Lint all packages |
| `pnpm -r typecheck` | Type-check all packages |
| `pnpm --filter @axi-office/core build` | Build one package |
| `pnpm changeset` | Create a new changeset |
| `pnpm changeset version` | Bump versions per changesets |
| `pnpm changeset publish` | Publish changed packages to npm |

## Conventions

- **ESM only** — all packages use `"type": "module"`.
- **Node >= 20** — required by `@modelcontextprotocol/sdk`.
- **Strict TypeScript** — extends root `tsconfig.json` which enables all strict flags.
- **TOON output** — use `@axi-office/core`'s `field`, `renderList`, `renderHelp` helpers; do not invent custom output formats.
- **4 fields max per list row** — TOON is token-efficient when rows are concise.
- **Biome** for linting and formatting — run `pnpm lint:fix` before committing.
- **One changeset per PR** — required for the release workflow to function.
