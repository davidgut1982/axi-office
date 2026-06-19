/**
 * Why: runAxiCli requires a `home` handler shown when the CLI is invoked with no
 * command; it should summarize the tool and its commands.
 * What: Returns a plain object describing the binary, which runAxiCli renders as TOON.
 * Test: Call homeCommand([]) and assert the result has bin === "excel-axi" and a
 * non-empty commands array.
 */

export function homeCommand(_args: string[]): Record<string, unknown> {
  return {
    bin: "excel-axi",
    description: "AXI CLI for Excel (via negokaz/excel-mcp-server)",
    commands: ["sheets", "read", "write", "create-table", "copy-sheet", "format-range"],
    help: "Run `excel-axi <command> --help` for details",
  };
}
