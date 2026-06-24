/**
 * Why: The GongRzhe PowerPoint MCP backend is stateful — it holds presentations in a
 * keyed dict and every subsequent call requires the presentation_id returned by create
 * or open. Every ppt-axi invocation is a fresh subprocess, so each command must be
 * self-contained: open (or create) → capture presentation_id → operate with id → save.
 * What: Exports four helpers — withOpenSave, withOpenReadonly, withCreateSave,
 * withCreateFromTemplateSave — that manage the lifecycle, plus a call() guard that
 * surfaces backend error responses as AxiError instead of silently succeeding.
 * Test: Mock getClient().callTool to return { presentation_id: "presentation_1" } for
 * create/open calls and { ok: true } for others. Assert presentation_id is threaded
 * through every subsequent call including save_presentation.
 */
import { AxiError } from "@axi-office/core";
import { getClient } from "./client.js";

type SessionFn = (client: ReturnType<typeof getClient>, presentationId: string) => Promise<unknown>;

/**
 * Why: Backend errors are returned as { error: "..." } with isError=false, so the CLI
 * would exit 0 silently without this guard.
 * What: Calls client.callTool and throws AxiError TOOL_ERROR if the result has a truthy
 * error property or success === false. Returns the result otherwise.
 */
export async function call(
	client: ReturnType<typeof getClient>,
	tool: string,
	args: Record<string, unknown>
): Promise<unknown> {
	const res = await client.callTool(tool, args);
	if (res !== null && typeof res === "object") {
		const r = res as Record<string, unknown>;
		if (r.error || r.success === false) {
			throw new AxiError(String(r.error ?? "backend error"), "TOOL_ERROR");
		}
	}
	return res;
}

/**
 * Why: Mutating commands must open, capture the presentation_id, operate, then save.
 * What: Calls open_presentation, extracts presentation_id from result, awaits
 * fn(client, presentationId), then calls save_presentation with the id.
 * Test: Assert callTool order: open_presentation → [fn calls] → save_presentation,
 * with presentation_id threaded into save args.
 */
export async function withOpenSave(filePath: string, fn: SessionFn): Promise<unknown> {
	const client = getClient();
	const openRes = await call(client, "open_presentation", { file_path: filePath });
	const presentationId = (openRes as Record<string, unknown>)?.presentation_id as
		| string
		| undefined;
	if (!presentationId) {
		throw new AxiError("open_presentation returned no presentation_id", "TOOL_ERROR");
	}
	const result = await fn(client, presentationId);
	await call(client, "save_presentation", {
		file_path: filePath,
		presentation_id: presentationId,
	});
	return result;
}

/**
 * Why: Read-only commands must open but must NOT save (to avoid dirty writes).
 * What: Calls open_presentation, extracts presentation_id, awaits fn(client, presentationId),
 * returns fn's result without saving.
 * Test: Assert callTool called with open_presentation but NOT save_presentation,
 * and that presentation_id is forwarded to fn.
 */
export async function withOpenReadonly(filePath: string, fn: SessionFn): Promise<unknown> {
	const client = getClient();
	const openRes = await call(client, "open_presentation", { file_path: filePath });
	const presentationId = (openRes as Record<string, unknown>)?.presentation_id as
		| string
		| undefined;
	if (!presentationId) {
		throw new AxiError("open_presentation returned no presentation_id", "TOOL_ERROR");
	}
	return fn(client, presentationId);
}

/**
 * Why: New presentations must be created in-memory, capturing the presentation_id,
 * before operating and then saving.
 * What: Calls create_presentation, extracts presentation_id, awaits fn(client, presentationId),
 * then saves to filePath with the id.
 * Test: Assert callTool order: create_presentation → [fn calls] → save_presentation,
 * with presentation_id in save args.
 */
export async function withCreateSave(filePath: string, fn: SessionFn): Promise<unknown> {
	const client = getClient();
	const createRes = await call(client, "create_presentation", {});
	const presentationId = (createRes as Record<string, unknown>)?.presentation_id as
		| string
		| undefined;
	if (!presentationId) {
		throw new AxiError("create_presentation returned no presentation_id", "TOOL_ERROR");
	}
	const result = await fn(client, presentationId);
	await call(client, "save_presentation", {
		file_path: filePath,
		presentation_id: presentationId,
	});
	return result;
}

/**
 * Why: Template-based creation needs the same lifecycle as withCreateSave but uses
 * create_presentation_from_template. Routing through this helper ensures the
 * open→operate→save contract is in one tested abstraction.
 * What: Calls create_presentation_from_template({template_path}), extracts presentation_id,
 * awaits fn(client, presentationId), then saves to filePath with the id.
 * Test: Assert callTool order: create_presentation_from_template → [fn calls] →
 * save_presentation, with presentation_id in save args.
 */
export async function withCreateFromTemplateSave(
	templatePath: string,
	filePath: string,
	fn: SessionFn
): Promise<unknown> {
	const client = getClient();
	const createRes = await call(client, "create_presentation_from_template", {
		template_path: templatePath,
	});
	const presentationId = (createRes as Record<string, unknown>)?.presentation_id as
		| string
		| undefined;
	if (!presentationId) {
		throw new AxiError(
			"create_presentation_from_template returned no presentation_id",
			"TOOL_ERROR"
		);
	}
	const result = await fn(client, presentationId);
	await call(client, "save_presentation", {
		file_path: filePath,
		presentation_id: presentationId,
	});
	return result;
}
