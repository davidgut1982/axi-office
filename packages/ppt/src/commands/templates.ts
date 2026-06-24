import { getClient } from "../client.js";
/**
 * Why: Users need to discover available slide templates before using --color-scheme
 * in add-slide or auto-generate; this maps `templates` to list_slide_templates.
 * What: Calls list_slide_templates directly on the client — no presentation open or
 * save required since this is a pure registry query. Routes through call() guard
 * to surface backend error responses as AxiError.
 * Test: Mock the client, call templatesCommand([]), assert callTool was invoked with
 * "list_slide_templates" and no other calls before or after.
 */
import { call } from "../session.js";

export async function templatesCommand(_args: string[]): Promise<unknown> {
	const client = getClient();
	return call(client, "list_slide_templates", {});
}
