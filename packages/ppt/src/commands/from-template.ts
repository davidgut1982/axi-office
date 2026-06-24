/**
 * Why: Templates allow consistent branding and layout across presentations;
 * this maps `from-template` to withCreateFromTemplateSave so the lifecycle
 * (create_from_template → operate → save) goes through the tested session abstraction.
 * What: Validates <file> and <template-path>, then calls withCreateFromTemplateSave
 * which issues create_presentation_from_template followed by save_presentation.
 * Test: Mock the client, call fromTemplateCommand(["/tmp/out.pptx", "/tmp/tmpl.pptx"]),
 * assert callTool called with "create_presentation_from_template" then "save_presentation".
 */
import { AxiError, parseFlags } from "@axi-office/core";
import { withCreateFromTemplateSave } from "../session.js";

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

	return withCreateFromTemplateSave(templatePath, file, async (_client) => {
		return { created: file, template: templatePath };
	});
}
