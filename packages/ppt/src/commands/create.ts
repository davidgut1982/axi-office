/**
 * Why: Creating a blank presentation is the entry point for building decks from scratch;
 * this maps `create` to the create_presentation + save_presentation MCP tools.
 * What: Validates the <file> positional then uses withCreateSave to create a blank deck
 * and immediately persist it to disk at <file>.
 * Test: Mock the client, call createCommand(["/tmp/x.pptx"]), assert callTool was
 * invoked first with "create_presentation" and then with "save_presentation".
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { withCreateSave } from "../session.js";

export async function createCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const file = positionals[0];
	if (!file) {
		throw new AxiError("file path is required", "VALIDATION_ERROR", [
			"ppt-axi create <file>",
			"",
			"Create a new blank PowerPoint presentation at <file>.",
			"Requires uv/uvx on PATH (curl -LsSf https://astral.sh/uv/install.sh | sh)",
		]);
	}

	return withCreateSave(file, async (_client) => {
		return { created: file };
	});
}
