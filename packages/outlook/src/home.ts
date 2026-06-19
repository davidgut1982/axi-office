/**
 * Why: runAxiCli requires a `home` handler shown when invoked with no command.
 * What: Returns a plain object describing the binary and its commands for TOON output.
 * Test: Call homeCommand([]) and assert bin === "outlook-axi" and a non-empty commands array.
 */

export function homeCommand(_args: string[]): Record<string, unknown> {
  return {
    bin: "outlook-axi",
    description: "AXI CLI for Outlook (via softeria/ms-365-mcp-server)",
    commands: [
      "mail-list",
      "mail-get",
      "mail-send",
      "folders",
      "cal-list",
      "cal-view",
      "cal-create",
      "contacts-list",
      "login",
    ],
    help: "Run `outlook-axi <command> --help` for details",
  };
}
