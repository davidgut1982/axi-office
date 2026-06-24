#!/usr/bin/env node
/**
 * Why: This is the published `ppt-axi` binary entry point; it wires the command
 * handlers into runAxiCli so the AXI runtime handles parsing, help, and TOON rendering.
 * What: Reads the package version, normalizes each MCP command result into an
 * AxiRenderable, and registers all ppt commands plus `setup hooks`.
 * Test: Run `node dist/bin/ppt-axi.js --help` and assert the top-level help prints;
 * run with no args and assert the home object prints.
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { AxiError, closeAllLiveClients, runAxiCli, setupHooksCommand } from "@axi-office/core";
import { addChartCommand } from "../commands/add-chart.js";
import { addImageCommand } from "../commands/add-image.js";
import { addShapeCommand } from "../commands/add-shape.js";
import { addSlideCommand } from "../commands/add-slide.js";
import { addTableCommand } from "../commands/add-table.js";
import { addTextCommand } from "../commands/add-text.js";
import { autoGenerateCommand } from "../commands/auto-generate.js";
import { bulletsCommand } from "../commands/bullets.js";
import { createCommand } from "../commands/create.js";
import { fromTemplateCommand } from "../commands/from-template.js";
import { infoCommand } from "../commands/info.js";
import { readSlideCommand } from "../commands/read-slide.js";
import { readCommand } from "../commands/read.js";
import { setPlaceholderCommand } from "../commands/set-placeholder.js";
import { setPropsCommand } from "../commands/set-props.js";
import { slideInfoCommand } from "../commands/slide-info.js";
import { templatesCommand } from "../commands/templates.js";
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
 * Why: MCP tool results are arbitrary JSON (objects, arrays, scalars, strings) but
 * runAxiCli commands must return a string or a plain object.
 * What: Wraps a handler, passing through strings and plain objects and boxing anything
 * else under a `result` key.
 * Test: render(async () => [1, 2]) resolves to { result: [1, 2] };
 * render(async () => "ok") resolves to "ok".
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
	description: "AXI CLI for PowerPoint (via GongRzhe/Office-PowerPoint-MCP-Server, run with uvx)",
	version: readVersion(),
	topLevelHelp: TOP_LEVEL_HELP,
	getCommandHelp: (command: string) => COMMAND_HELP[command] ?? null,
	home: (args: string[]) => homeCommand(args),
	commands: {
		create: render(createCommand),
		info: render(infoCommand),
		"set-props": render(setPropsCommand),
		"from-template": render(fromTemplateCommand),
		"add-slide": render(addSlideCommand),
		"slide-info": render(slideInfoCommand),
		read: render(readCommand),
		"read-slide": render(readSlideCommand),
		"set-placeholder": render(setPlaceholderCommand),
		bullets: render(bulletsCommand),
		"add-text": render(addTextCommand),
		"add-image": render(addImageCommand),
		"add-table": render(addTableCommand),
		"add-shape": render(addShapeCommand),
		"add-chart": render(addChartCommand),
		templates: render(templatesCommand),
		"auto-generate": render(autoGenerateCommand),
		setup: async (args: string[]) => {
			if (args[0] === "hooks") {
				return setupHooksCommand("ppt-axi", ["ppt-axi"]);
			}
			throw new AxiError("setup subcommand is required", "VALIDATION_ERROR", [
				"ppt-axi setup hooks    — Install session-start hooks",
			]);
		},
	},
});

// Await graceful teardown of all live MCP child processes before exiting so
// the Python server has a chance to flush stderr and exit cleanly.
await closeAllLiveClients();
process.exit(process.exitCode ?? 0);
