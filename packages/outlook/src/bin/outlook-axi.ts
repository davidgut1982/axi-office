#!/usr/bin/env node
/**
 * Why: This is the published `outlook-axi` binary entry point; it wires the command
 * handlers into runAxiCli so the AXI runtime handles parsing, help, and TOON rendering.
 * What: Reads the package version, normalizes each command result into an AxiRenderable,
 * and registers all outlook commands plus `setup hooks`.
 * Test: Run `node dist/bin/outlook-axi.js --help` and assert the top-level help prints; run
 * with no args and assert the home object prints.
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { AxiError, runAxiCli, setupHooksCommand } from "@axi-office/core";
import { calCreateCommand, calListCommand, calViewCommand } from "../commands/calendar.js";
import { contactsListCommand } from "../commands/contacts.js";
import { loginCommand } from "../commands/login.js";
import {
	foldersCommand,
	mailGetCommand,
	mailListCommand,
	mailSendCommand,
} from "../commands/mail.js";
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
 * Why: MCP tool results are arbitrary JSON but runAxiCli commands must return a string or a
 * plain object.
 * What: Wraps a handler, passing through strings and plain objects and boxing anything else.
 * Test: render(async () => ["a"])([]) resolves to { result: ["a"] }.
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
	description: "AXI CLI for Outlook (via softeria/ms-365-mcp-server)",
	version: readVersion(),
	topLevelHelp: TOP_LEVEL_HELP,
	getCommandHelp: (command: string) => COMMAND_HELP[command] ?? null,
	home: (args: string[]) => homeCommand(args),
	commands: {
		"mail-list": render(mailListCommand),
		"mail-get": render(mailGetCommand),
		"mail-send": render(mailSendCommand),
		folders: render(foldersCommand),
		"cal-list": render(calListCommand),
		"cal-view": render(calViewCommand),
		"cal-create": render(calCreateCommand),
		"contacts-list": render(contactsListCommand),
		login: render(loginCommand),
		setup: async (args: string[]) => {
			if (args[0] === "hooks") {
				return setupHooksCommand("outlook-axi", ["outlook-axi"]);
			}
			throw new AxiError("setup subcommand is required", "VALIDATION_ERROR", [
				"outlook-axi setup hooks    — Install session-start hooks",
			]);
		},
	},
});

// Force exit so the MCP subprocess does not keep Node alive after the command
// completes. The "exit" event fires synchronously, triggering _syncKillChild()
// on every live McpStdioClient so the child is SIGKILL'd before Node exits.
process.exit(process.exitCode ?? 0);
