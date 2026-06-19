#!/usr/bin/env node
/**
 * Why: This is the published `excel-axi` binary entry point; it wires the command
 * handlers into runAxiCli so the AXI runtime handles parsing, help, and TOON rendering.
 * What: Reads the package version, normalizes each MCP command result into an
 * AxiRenderable, and registers all excel commands plus `setup hooks`.
 * Test: Run `node dist/bin/excel-axi.js --help` and assert the top-level help prints;
 * run with no args and assert the home object prints.
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { AxiError, runAxiCli, setupHooksCommand } from "@axi-office/core";
import { copySheetCommand } from "../commands/copy-sheet.js";
import { createTableCommand } from "../commands/create-table.js";
import { formatRangeCommand } from "../commands/format-range.js";
import { readCommand } from "../commands/read.js";
import { sheetsCommand } from "../commands/sheets.js";
import { writeCommand } from "../commands/write.js";
import { COMMAND_HELP, TOP_LEVEL_HELP } from "../help.js";
import { homeCommand } from "../home.js";

function readVersion(): string {
	try {
		const here = dirname(fileURLToPath(import.meta.url));
		const pkg = JSON.parse(
			readFileSync(join(here, "../../package.json"), "utf8"),
		);
		return typeof pkg.version === "string" ? pkg.version : "0.0.0";
	} catch {
		return "0.0.0";
	}
}

/**
 * Why: MCP tool results are arbitrary JSON (objects, arrays, scalars, strings) but
 * runAxiCli commands must return a string or a plain object.
 * What: Wraps a handler, passing through strings and plain objects and boxing anything
 * else under a `result` key.
 * Test: render(async () => [1, 2]) resolves to { result: [1, 2] };
 * render(async () => "ok") resolves to "ok".
 */
function render(
	handler: (args: string[]) => Promise<unknown>,
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
	description: "AXI CLI for Excel (via negokaz/excel-mcp-server)",
	version: readVersion(),
	topLevelHelp: TOP_LEVEL_HELP,
	getCommandHelp: (command: string) => COMMAND_HELP[command] ?? null,
	home: (args: string[]) => homeCommand(args),
	commands: {
		sheets: render(sheetsCommand),
		read: render(readCommand),
		write: render(writeCommand),
		"create-table": render(createTableCommand),
		"copy-sheet": render(copySheetCommand),
		"format-range": render(formatRangeCommand),
		setup: async (args: string[]) => {
			if (args[0] === "hooks") {
				return setupHooksCommand("excel-axi", ["excel-axi"]);
			}
			throw new AxiError("setup subcommand is required", "VALIDATION_ERROR", [
				"excel-axi setup hooks    — Install session-start hooks",
			]);
		},
	},
});
