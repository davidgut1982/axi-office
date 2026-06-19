/**
 * Why: Calendar read/write is the second Outlook surface; these handlers map the cal-*
 * commands to the ms-365-mcp-server calendar tools.
 * What: calListCommand -> list-calendar-events, calViewCommand -> get-calendar-view,
 * calCreateCommand -> create-calendar-event, with flag/positional validation.
 * calViewCommand accepts full ISO 8601 datetimes or bare YYYY-MM-DD dates; bare dates are
 * expanded to T00:00:00Z (start) and T23:59:59Z (end) so the server always receives a
 * valid datetime string.
 * Test: Mock the client, call calViewCommand(["2026-01-01","2026-01-31"]); assert callTool
 * was invoked with "get-calendar-view" and startDateTime "2026-01-01T00:00:00Z" and
 * endDateTime "2026-01-31T23:59:59Z".
 */
import { AxiError, parseFlags, parseLimit } from "@axi-office/core";
import { getClient } from "../client.js";

const BARE_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
// ISO 8601 datetime: date + time component (T separator required)
const ISO_DATETIME_RE = /^\d{4}-\d{2}-\d{2}T[\d:.]+([Zz]|[+-]\d{2}:\d{2})?$/;

/**
 * Why: The ms-365-mcp-server get-calendar-view tool requires full ISO 8601 datetime
 * strings; this helper normalises bare dates so callers can pass just a date.
 * What: If input is a bare YYYY-MM-DD date, appends a time suffix; if it already looks
 * like an ISO datetime it is returned as-is; otherwise throws AxiError.
 * Test: normalizeDateTime("2026-01-01", "start") → "2026-01-01T00:00:00Z";
 *       normalizeDateTime("2026-01-31", "end")   → "2026-01-31T23:59:59Z";
 *       normalizeDateTime("2026-01-01T09:00:00Z", "start") → "2026-01-01T09:00:00Z".
 */
function normalizeDateTime(value: string, role: "start" | "end"): string {
	if (ISO_DATETIME_RE.test(value)) {
		return value;
	}
	if (BARE_DATE_RE.test(value)) {
		const suffix = role === "start" ? "T00:00:00Z" : "T23:59:59Z";
		return `${value}${suffix}`;
	}
	throw new AxiError(`"${value}" is not a valid date or datetime for ${role}`, "VALIDATION_ERROR", [
		"Accepted formats:",
		"  YYYY-MM-DD           (bare date, e.g. 2026-01-01)",
		"  YYYY-MM-DDTHH:MM:SSZ (ISO 8601 datetime, e.g. 2026-01-01T09:00:00Z)",
	]);
}

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
			"",
			"<start> and <end> may be bare dates (YYYY-MM-DD) or ISO 8601 datetimes.",
			"Bare dates are expanded: start → T00:00:00Z, end → T23:59:59Z.",
		]);
	}
	const startDateTime = normalizeDateTime(start, "start");
	const endDateTime = normalizeDateTime(end, "end");
	return getClient().callTool("get-calendar-view", {
		startDateTime,
		endDateTime,
	});
}

export async function calCreateCommand(args: string[]): Promise<unknown> {
	const { positionals, flags } = parseFlags(args);
	const [subject, start, end] = positionals;
	if (!subject || !start || !end) {
		throw new AxiError("subject, start and end are required", "VALIDATION_ERROR", [
			"outlook-axi cal-create <subject> <start> <end> [--attendees a@x,b@y]",
		]);
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
