/**
 * Why: Path sandbox correctness is a security invariant; these tests prove that
 * resolveInBase allows in-base paths and rejects traversal before any FS op.
 * What: Unit tests for resolveInBase covering: in-base file, traversal via ../,
 * absolute path outside base, and undefined base (no sandbox).
 * Test: This file is the test.
 */
import { resolve } from "node:path";
import { AxiError } from "@axi-office/core";
import { describe, expect, it } from "vitest";
import { resolveInBase } from "../src/paths.js";

describe("resolveInBase", () => {
	const base = "/safe/dir";

	it("passes a file directly inside base", () => {
		const result = resolveInBase(base, "/safe/dir/file.docx");
		expect(result).toBe(resolve("/safe/dir/file.docx"));
	});

	it("passes a nested file inside base", () => {
		const result = resolveInBase(base, "/safe/dir/sub/file.docx");
		expect(result).toBe(resolve("/safe/dir/sub/file.docx"));
	});

	it("rejects traversal via ../", () => {
		expect(() => resolveInBase(base, "/safe/dir/../etc/passwd")).toThrow(
			AxiError,
		);
		try {
			resolveInBase(base, "/safe/dir/../etc/passwd");
		} catch (e) {
			expect((e as AxiError).code).toBe("SECURITY_ERROR");
		}
	});

	it("rejects absolute path outside base", () => {
		expect(() => resolveInBase(base, "/etc/passwd")).toThrow(AxiError);
		try {
			resolveInBase(base, "/etc/passwd");
		} catch (e) {
			expect((e as AxiError).code).toBe("SECURITY_ERROR");
		}
	});

	it("rejects a path that is a prefix match but not a subpath (e.g. /safe/dirother)", () => {
		expect(() => resolveInBase(base, "/safe/dirother/file.docx")).toThrow(
			AxiError,
		);
	});

	it("returns resolved path without sandbox when baseDir is undefined", () => {
		const result = resolveInBase(undefined, "relative/file.docx");
		expect(result).toBe(resolve("relative/file.docx"));
	});
});
