/**
 * Why: runAxiCli requires a `home` handler shown when invoked with no command.
 * What: Returns a plain object describing the binary and its commands for TOON output.
 * Test: Call homeCommand([]) and assert bin === "word-axi" and a non-empty commands array.
 */

export function homeCommand(_args: string[]): Record<string, unknown> {
  return {
    bin: "word-axi",
    description: "AXI CLI for Word (.docx)",
    commands: ["create", "from-markdown", "read", "info", "patch"],
    help: "Run `word-axi <command> --help` for details",
  };
}
