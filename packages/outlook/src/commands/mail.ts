/**
 * Why: Mail is the core Outlook surface; these handlers map the mail-* commands to the
 * ms-365-mcp-server mail tools.
 * What: mailListCommand -> list-mail-messages, mailGetCommand -> get-mail-message,
 * mailSendCommand -> send-mail, foldersCommand -> list-mail-folders, with flag validation.
 * Test: Mock the client, call mailListCommand(["--limit","5"]); assert callTool was invoked
 * with "list-mail-messages" and a top of 5.
 */
import { AxiError, parseFlags, parseLimit } from "@axi-office/core";
import { getClient } from "../client.js";

export async function mailListCommand(args: string[]): Promise<unknown> {
	const { flags } = parseFlags(args);
	const toolArgs: Record<string, unknown> = {};
	if (flags.limit !== undefined) {
		toolArgs.$top = parseLimit(flags.limit, 25, 100);
	}
	if (typeof flags.folder === "string") {
		toolArgs.mailFolder = flags.folder;
	}
	return getClient().callTool("list-mail-messages", toolArgs);
}

export async function mailGetCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const id = positionals[0];
	if (!id) {
		throw new AxiError("message id is required", "VALIDATION_ERROR", [
			"outlook-axi mail-get <id>",
		]);
	}
	return getClient().callTool("get-mail-message", { messageId: id });
}

export async function mailSendCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [to, subject, body] = positionals;
	if (!to || !subject || !body) {
		throw new AxiError(
			"to, subject and body are required",
			"VALIDATION_ERROR",
			["outlook-axi mail-send <to> <subject> <body>"],
		);
	}
	return getClient().callTool("send-mail", {
		message: {
			subject,
			body: { contentType: "Text", content: body },
			toRecipients: [{ emailAddress: { address: to } }],
		},
	});
}

export async function foldersCommand(_args: string[]): Promise<unknown> {
	return getClient().callTool("list-mail-folders", {});
}
