/**
 * Why: Calendar read/write is the second Outlook surface; these handlers map the cal-*
 * commands to the ms-365-mcp-server calendar tools.
 * What: calListCommand -> list-calendar-events, calViewCommand -> get-calendar-view,
 * calCreateCommand -> create-calendar-event, with flag/positional validation.
 * Test: Mock the client, call calViewCommand(["2026-01-01","2026-01-31"]); assert callTool
 * was invoked with "get-calendar-view" and the start/end datetimes.
 */
import { AxiError, parseFlags, parseLimit } from "@axi-office/core";
import { getClient } from "../client.js";

export async function calListCommand(args: string[]): Promise<unknown> {
	const { flags } = parseFlags(args);
	const toolArgs: Record<string, unknown> = {};
	if (flags.limit !== undefined) {
		toolArgs.$top = parseLimit(flags.limit, 25, 100);
	}
	return getClient().callTool("list-calendar-events", toolArgs);
}

export async function calViewCommand(args: string[]): Promise<unknown> {
	const { positionals } = parseFlags(args);
	const [start, end] = positionals;
	if (!start || !end) {
		throw new AxiError("start and end are required", "VALIDATION_ERROR", [
			"outlook-axi cal-view <start> <end>",
		]);
	}
	return getClient().callTool("get-calendar-view", {
		startDateTime: start,
		endDateTime: end,
	});
}

export async function calCreateCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args);
	const [subject, start, end] = positionals;
	if (!subject || !start || !end) {
		throw new AxiError(
			"subject, start and end are required",
			"VALIDATION_ERROR",
			["outlook-axi cal-create <subject> <start> <end> [--attendees a@x,b@y]"],
		);
	}

	const event: Record<string, unknown> = {
		subject,
		start: { dateTime: start, timeZone: "UTC" },
		end: { dateTime: end, timeZone: "UTC" },
	};

	if (typeof flags.attendees === "string" && flags.attendees.length > 0) {
		event.attendees = flags.attendees
			.split(",")
			.map((a) => a.trim())
			.filter((a) => a.length > 0)
			.map((address) => ({ emailAddress: { address }, type: "required" }));
	}

	return getClient().callTool("create-calendar-event", event);
}
