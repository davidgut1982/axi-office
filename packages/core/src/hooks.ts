/**
 * Why: Every @axi-office CLI binary needs to expose a "setup hooks" subcommand that
 * installs the axi-sdk-js SessionStart hook into Claude Code / Codex settings. This
 * helper avoids duplicating the installSessionStartHooks() call in every binary.
 *
 * What: Exports setupHooksCommand() which calls installSessionStartHooks() with the
 * given marker and binaryNames, then returns a plain result object for runAxiCli().
 *
 * Test: Mock installSessionStartHooks (import it in your test, then vi.mock it);
 * call setupHooksCommand("excel-axi", ["excel-axi"]); assert installSessionStartHooks
 * was called with { marker: "excel-axi", binaryNames: ["excel-axi"] } and the
 * returned object has setup: "hooks installed or already up to date".
 */
import { installSessionStartHooks } from "axi-sdk-js";

/**
 * Why: Packages the session hook installation result into a consistent plain object
 * that runAxiCli() can render as TOON without any per-package boilerplate.
 * What: Calls installSessionStartHooks({ marker, binaryNames }) and returns a success
 * record; errors from the hook installation are surfaced via the onError callback and
 * collected into the warnings field.
 * Test: Call setupHooksCommand("my-axi", ["my-axi"]) and assert the returned object
 * contains { setup: "hooks installed or already up to date" }.
 */
export function setupHooksCommand(marker: string, binaryNames: string[]): Record<string, unknown> {
  const warnings: string[] = [];

  installSessionStartHooks({
    marker,
    binaryNames,
    onError: (msg: string) => warnings.push(msg),
  });

  const result: Record<string, unknown> = {
    setup: "hooks installed or already up to date",
    marker,
  };

  if (warnings.length > 0) {
    result.warnings = warnings;
  }

  return result;
}
