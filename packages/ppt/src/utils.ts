/**
 * Why: parseColor is used in add-text and add-shape with the same logic but previously
 * existed as duplicate local functions with diverged signatures. Centralizing it
 * prevents drift and makes the constraint (0-255 per channel) a single source of truth.
 * What: Parses a "R,G,B" string into a typed [number, number, number] tuple, throwing
 * AxiError VALIDATION_ERROR with a hint showing the flag name on any parse failure.
 * Test: parseColor("255,0,128", "--color") → [255, 0, 128];
 * parseColor("notacolor", "--color") → throws AxiError VALIDATION_ERROR.
 */
import { AxiError } from "@axi-office/core";

export function parseColor(colorStr: string, flagName: string): [number, number, number] {
	const parts = colorStr.split(",").map((s) => Number.parseInt(s.trim(), 10));
	if (parts.length !== 3 || parts.some((n) => !Number.isFinite(n) || n < 0 || n > 255)) {
		throw new AxiError(
			`${flagName} "${colorStr}" is not a valid R,G,B triplet`,
			"VALIDATION_ERROR",
			[`Example: ${flagName} 255,0,0`]
		);
	}
	return parts as [number, number, number];
}
