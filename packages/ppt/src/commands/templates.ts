/**
 * Why: Users need to discover available slide templates before using --color-scheme
 * in add-slide or auto-generate; this maps `templates` to list_slide_templates.
 * What: Calls list_slide_templates directly on the client — no presentation open or
 * save required since this is a pure registry query.
 * Test: Mock the client, call templatesCommand([]), assert callTool was invoked with
 * "list_slide_templates" and no other calls before or after.
 */
import { getClient } from "../client.js";

export async function templatesCommand(_args: string[]): Promise<unknown> {
	return getClient().callTool("list_slide_templates", {});
}
