/**
 * Why: Document metadata (title, author, subject, keywords, comments) is important
 * for accessibility and enterprise workflows; this maps `set-props` to set_core_properties.
 * What: Validates <file>, requires at least one property flag, opens with save, and
 * calls set_core_properties with only the provided flags (no undefined values sent).
 * Test: Mock the client, call setPropsCommand(["/tmp/x.pptx", "--title", "My Deck"]),
 * assert callTool was invoked with set_core_properties and { title: "My Deck" } only.
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { withOpenSave } from "../session.js";

export async function setPropsCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args);
	const file = positionals[0];
	if (!file) {
		throw new AxiError("file path is required", "VALIDATION_ERROR", [
			"ppt-axi set-props <file> [--title T] [--author A] [--subject S] [--keywords K] [--comments C]",
		]);
	}

	const propArgs: Record<string, unknown> = {};
	if (typeof flags.title === "string") propArgs.title = flags.title;
	if (typeof flags.author === "string") propArgs.author = flags.author;
	if (typeof flags.subject === "string") propArgs.subject = flags.subject;
	if (typeof flags.keywords === "string") propArgs.keywords = flags.keywords;
	if (typeof flags.comments === "string") propArgs.comments = flags.comments;

	if (Object.keys(propArgs).length === 0) {
		throw new AxiError("at least one property flag is required", "VALIDATION_ERROR", [
			"ppt-axi set-props <file> [--title T] [--author A] [--subject S] [--keywords K] [--comments C]",
		]);
	}

	return withOpenSave(file, async (client) => {
		return client.callTool("set_core_properties", propArgs);
	});
}
