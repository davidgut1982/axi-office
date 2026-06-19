/**
 * Why: Agents processing untrusted specs or markdown may supply attacker-controlled
 * file paths; this helper enforces an opt-in base-directory sandbox so path traversal
 * (../../etc/passwd) is rejected at the CLI boundary before any FS operation.
 * What: resolveInBase resolves a path against base and throws AxiError("SECURITY_ERROR")
 * if the result escapes the base directory. When base is undefined the path is returned
 * as-is (opt-in; default behavior is unchanged).
 * Test: resolveInBase("/safe/dir", "/safe/dir/file.docx") returns resolved path;
 * resolveInBase("/safe/dir", "../etc/passwd") throws AxiError with SECURITY_ERROR;
 * resolveInBase(undefined, "any/path") returns resolved path without sandbox.
 */
import { resolve } from "node:path";
import { AxiError } from "@axi-office/core";

/**
 * Resolves `filePath` and, when `baseDir` is provided, asserts it stays within
 * `baseDir`. Throws AxiError("SECURITY_ERROR") on traversal.
 *
 * @internal
 */
export function resolveInBase(
	baseDir: string | undefined,
	filePath: string,
): string {
	const resolved = resolve(filePath);
	if (baseDir === undefined) return resolved;
	const base = resolve(baseDir);
	if (!resolved.startsWith(`${base}/`) && resolved !== base) {
		throw new AxiError(
			`path "${filePath}" escapes base dir "${base}"`,
			"SECURITY_ERROR",
			[`Resolved path: ${resolved}`, `Base dir:      ${base}`],
		);
	}
	return resolved;
}
