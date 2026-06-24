/**
 * Why: The GongRzhe PowerPoint MCP backend is stateful — it holds the "current
 * presentation" in memory and only persists it on save_presentation. Every ppt-axi
 * invocation is a fresh subprocess, so each command must be self-contained:
 * open (or create) → operate → save, all over the single McpStdioClient connection.
 * What: Exports four helpers — withOpenSave, withOpenReadonly, withCreateSave,
 * withCreateFromTemplateSave — that
 * manage the open/create → operate → optional-save lifecycle so command handlers only
 * implement the "operate" step.
 * Test: Mock getClient().callTool; call withOpenSave("/tmp/x.pptx", fn) and assert
 * open_presentation was called first, then the fn result, then save_presentation last.
 * For withOpenReadonly assert save_presentation is NOT called.
 */
import { getClient } from "./client.js";

/**
 * Why: Mutating commands must open, operate, then save in order.
 * What: Calls open_presentation, awaits fn(client), then calls save_presentation.
 * Returns the result from fn.
 * Test: Assert callTool order: open_presentation → [fn calls] → save_presentation.
 */
export async function withOpenSave(
	filePath: string,
	fn: (client: ReturnType<typeof getClient>) => Promise<unknown>
): Promise<unknown> {
	const client = getClient();
	await client.callTool("open_presentation", { file_path: filePath });
	const result = await fn(client);
	await client.callTool("save_presentation", { file_path: filePath });
	return result;
}

/**
 * Why: Read-only commands must open but must NOT save (to avoid dirty writes).
 * What: Calls open_presentation, awaits fn(client), returns fn's result without saving.
 * Test: Assert callTool called with open_presentation but NOT save_presentation.
 */
export async function withOpenReadonly(
	filePath: string,
	fn: (client: ReturnType<typeof getClient>) => Promise<unknown>
): Promise<unknown> {
	const client = getClient();
	await client.callTool("open_presentation", { file_path: filePath });
	return fn(client);
}

/**
 * Why: New presentations must be created in-memory before operating and then saved.
 * What: Calls create_presentation (no args, uses backend default), awaits fn(client),
 * then saves to filePath.
 * Test: Assert callTool order: create_presentation → [fn calls] → save_presentation.
 */
export async function withCreateSave(
	filePath: string,
	fn: (client: ReturnType<typeof getClient>) => Promise<unknown>
): Promise<unknown> {
	const client = getClient();
	await client.callTool("create_presentation", {});
	const result = await fn(client);
	await client.callTool("save_presentation", { file_path: filePath });
	return result;
}

/**
 * Why: Template-based creation needs the same lifecycle discipline as withCreateSave
 * but uses create_presentation_from_template instead of create_presentation. Routing
 * from-template through this helper ensures the open→operate→save contract is in
 * one tested abstraction rather than raw callTool calls in the command handler.
 * What: Calls create_presentation_from_template({template_path}), awaits fn(client),
 * then saves to filePath.
 * Test: Assert callTool order: create_presentation_from_template → [fn calls] → save_presentation.
 */
export async function withCreateFromTemplateSave(
	templatePath: string,
	filePath: string,
	fn: (client: ReturnType<typeof getClient>) => Promise<unknown>
): Promise<unknown> {
	const client = getClient();
	await client.callTool("create_presentation_from_template", { template_path: templatePath });
	const result = await fn(client);
	await client.callTool("save_presentation", { file_path: filePath });
	return result;
}
