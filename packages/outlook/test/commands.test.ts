/**
 * Why: Each command must call the correct ms-365-mcp-server tool with correctly shaped
 * arguments; these tests lock that wire contract so a refactor cannot silently break it.
 * What: Mocks the client factory and asserts tool name + arguments per command, plus
 * validation errors and the offline login guidance.
 * Test: This file is the test.
 */
import { AxiError } from "@axi-office/core";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mockCallTool = vi.fn();

vi.mock("../src/client.js", () => ({
	getClient: () => ({ callTool: mockCallTool, close: vi.fn() }),
}));

import {
	calCreateCommand,
	calListCommand,
	calViewCommand,
} from "../src/commands/calendar.js";
import { contactsListCommand } from "../src/commands/contacts.js";
import { loginCommand } from "../src/commands/login.js";
import {
	foldersCommand,
	mailGetCommand,
	mailListCommand,
	mailSendCommand,
} from "../src/commands/mail.js";

describe("outlook mail commands", () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCallTool.mockResolvedValue({ ok: true });
	});

	it("mail-list calls list-mail-messages with $top and folder", async () => {
		await mailListCommand(["--limit", "5", "--folder", "Inbox"]);
		expect(mockCallTool).toHaveBeenCalledWith("list-mail-messages", {
			$top: 5,
			mailFolder: "Inbox",
		});
	});

	it("mail-get calls get-mail-message", async () => {
		await mailGetCommand(["abc123"]);
		expect(mockCallTool).toHaveBeenCalledWith("get-mail-message", {
			messageId: "abc123",
		});
	});

	it("mail-get requires an id", async () => {
		await expect(mailGetCommand([])).rejects.toBeInstanceOf(AxiError);
	});

	it("mail-send shapes the Graph message payload", async () => {
		await mailSendCommand(["a@x.com", "Hi", "Body text"]);
		expect(mockCallTool).toHaveBeenCalledWith("send-mail", {
			message: {
				subject: "Hi",
				body: { contentType: "Text", content: "Body text" },
				toRecipients: [{ emailAddress: { address: "a@x.com" } }],
			},
		});
	});

	it("mail-send requires to, subject, body", async () => {
		await expect(mailSendCommand(["a@x.com"])).rejects.toBeInstanceOf(AxiError);
	});

	it("folders calls list-mail-folders", async () => {
		await foldersCommand([]);
		expect(mockCallTool).toHaveBeenCalledWith("list-mail-folders", {});
	});
});

describe("outlook calendar commands", () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCallTool.mockResolvedValue({ ok: true });
	});

	it("cal-list calls list-calendar-events", async () => {
		await calListCommand([]);
		expect(mockCallTool).toHaveBeenCalledWith("list-calendar-events", {});
	});

	it("cal-view calls get-calendar-view with range", async () => {
		await calViewCommand(["2026-01-01T00:00:00Z", "2026-01-31T00:00:00Z"]);
		expect(mockCallTool).toHaveBeenCalledWith("get-calendar-view", {
			startDateTime: "2026-01-01T00:00:00Z",
			endDateTime: "2026-01-31T00:00:00Z",
		});
	});

	it("cal-view requires start and end", async () => {
		await expect(calViewCommand(["only-start"])).rejects.toBeInstanceOf(
			AxiError,
		);
	});

	it("cal-create parses attendees", async () => {
		await calCreateCommand([
			"Standup",
			"2026-02-01T09:00:00",
			"2026-02-01T09:30:00",
			"--attendees",
			"a@x.com,b@y.com",
		]);
		expect(mockCallTool).toHaveBeenCalledWith("create-calendar-event", {
			subject: "Standup",
			start: { dateTime: "2026-02-01T09:00:00", timeZone: "UTC" },
			end: { dateTime: "2026-02-01T09:30:00", timeZone: "UTC" },
			attendees: [
				{ emailAddress: { address: "a@x.com" }, type: "required" },
				{ emailAddress: { address: "b@y.com" }, type: "required" },
			],
		});
	});
});

describe("outlook contacts + login", () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockCallTool.mockResolvedValue({ ok: true });
	});

	it("contacts-list calls list-outlook-contacts with $top", async () => {
		await contactsListCommand(["--limit", "10"]);
		expect(mockCallTool).toHaveBeenCalledWith("list-outlook-contacts", {
			$top: 10,
		});
	});

	it("login returns guidance without calling the MCP server", async () => {
		const result = (await loginCommand([])) as { steps: string[] };
		expect(mockCallTool).not.toHaveBeenCalled();
		expect(result.steps.join(" ")).toContain("--login");
	});
});
