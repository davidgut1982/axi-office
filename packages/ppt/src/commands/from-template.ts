/**
 * Why: Templates allow consistent branding and layout across presentations;
 * this maps `from-template` to create_presentation_from_template + save_presentation.
 * What: Validates <file> and <template-path>, calls create_presentation_from_template
 * with the template path, then saves the result to <file>.
 * Test: Mock the client, call fromTemplateCommand(["/tmp/out.pptx", "/tmp/tmpl.pptx"]),
 * assert callTool called with "create_presentation_from_template" then "save_presentation".
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { getClient } from "../client.js";

export async function fromTemplateCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [file, templatePath] = positionals;
	if (!file || !templatePath) {
		throw new AxiError("file and template-path are required", "VALIDATION_ERROR", [
			"ppt-axi from-template <file> <template-path>",
			"",
			"Create a presentation at <file> using <template-path> as the template.",
		]);
	}

	const client = getClient();
	const result = await client.callTool("create_presentation_from_template", {
		template_path: templatePath,
	});
	await client.callTool("save_presentation", { file_path: file });
	return result;
}
