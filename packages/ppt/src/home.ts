/**
 * Why: runAxiCli requires a `home` handler shown when the CLI is invoked with no
 * command; it should summarize the tool and its commands.
 * What: Returns a plain object describing the binary, which runAxiCli renders as TOON.
 * Test: Call homeCommand([]) and assert the result has bin === "ppt-axi" and a
 * non-empty commands array.
 */

export function homeCommand(_args: string[]): Record<string, unknown> {
	return {
		bin: "ppt-axi",
		description: "AXI CLI for PowerPoint (via GongRzhe/Office-PowerPoint-MCP-Server, run with uvx)",
		commands: [
			"create",
			"info",
			"set-props",
			"from-template",
			"add-slide",
			"slide-info",
			"read",
			"read-slide",
			"set-placeholder",
			"bullets",
			"add-text",
			"add-image",
			"add-table",
			"add-shape",
			"add-chart",
			"templates",
			"auto-generate",
		],
		help: "Run `ppt-axi <command> --help` for details",
	};
}
