#!/usr/bin/env node
/**
 * Why: This is the published `word-axi` binary entry point; it wires the command handlers
 * into runAxiCli so the AXI runtime handles parsing, help, and TOON rendering.
 * What: Reads the package version, normalizes each command result into an AxiRenderable,
 * and registers all word commands plus `setup hooks`.
 * Test: Run `node dist/bin/word-axi.js --help` and assert the top-level help prints; run
 * with no args and assert the home object prints.
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { AxiError, runAxiCli, setupHooksCommand } from "@axi-office/core";
import { createCommand } from "../commands/create.js";
import { fromMarkdownCommand } from "../commands/from-markdown.js";
import { infoCommand } from "../commands/info.js";
import { patchCommand } from "../commands/patch.js";
import { readCommand } from "../commands/read.js";
import { COMMAND_HELP, TOP_LEVEL_HELP } from "../help.js";
import { homeCommand } from "../home.js";

function readVersion(): string {
  try {
    const here = dirname(fileURLToPath(import.meta.url));
    const pkg = JSON.parse(readFileSync(join(here, "../../package.json"), "utf8"));
    return typeof pkg.version === "string" ? pkg.version : "0.0.0";
  } catch {
    return "0.0.0";
  }
}

/**
 * Why: Command handlers return plain result objects; runAxiCli accepts string or object,
 * so we pass objects through and box anything else.
 * What: Wraps a handler and normalizes its result into an AxiRenderable.
 * Test: render(async () => ({ a: 1 }))([]) resolves to { a: 1 }.
 */
function render(
  handler: (args: string[]) => Promise<unknown>
): (args: string[]) => Promise<string | Record<string, unknown>> {
  return async (args: string[]) => {
    const out = await handler(args);
    if (typeof out === "string") return out;
    if (out !== null && typeof out === "object" && !Array.isArray(out)) {
      return out as Record<string, unknown>;
    }
    return { result: out };
  };
}

await runAxiCli({
  description: "AXI CLI for Word (.docx)",
  version: readVersion(),
  topLevelHelp: TOP_LEVEL_HELP,
  getCommandHelp: (command: string) => COMMAND_HELP[command] ?? null,
  home: (args: string[]) => homeCommand(args),
  commands: {
    create: render(createCommand),
    "from-markdown": render(fromMarkdownCommand),
    read: render(readCommand),
    info: render(infoCommand),
    patch: render(patchCommand),
    setup: async (args: string[]) => {
      if (args[0] === "hooks") {
        return setupHooksCommand("word-axi", ["word-axi"]);
      }
      throw new AxiError("setup subcommand is required", "VALIDATION_ERROR", [
        "word-axi setup hooks    — Install session-start hooks",
      ]);
    },
  },
});
