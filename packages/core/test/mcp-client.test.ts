/**
 * Unit tests for McpStdioClient and McpClientError.
 *
 * We mock @modelcontextprotocol/sdk to avoid spawning real subprocesses in CI.
 * The tests verify argument shaping and lifecycle behavior.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";
import { McpClientError, McpStdioClient } from "../src/mcp-client.js";

// ---------------------------------------------------------------------------
// Mock @modelcontextprotocol/sdk
// ---------------------------------------------------------------------------

const mockCallTool = vi.fn();
const mockConnect = vi.fn();
const mockClose = vi.fn();

vi.mock("@modelcontextprotocol/sdk/client/index.js", () => ({
	Client: vi.fn().mockImplementation(() => ({
		connect: mockConnect,
		callTool: mockCallTool,
		close: mockClose,
	})),
}));

const mockTransportConstructor = vi.fn();

vi.mock("@modelcontextprotocol/sdk/client/stdio.js", () => ({
	StdioClientTransport: vi.fn().mockImplementation((opts: unknown) => {
		mockTransportConstructor(opts);
		return {};
	}),
}));

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("McpStdioClient", () => {
	beforeEach(() => {
		vi.clearAllMocks();
		mockConnect.mockResolvedValue(undefined);
		mockClose.mockResolvedValue(undefined);
	});

	it("lazily connects on first callTool()", async () => {
		mockCallTool.mockResolvedValue({
			content: [{ type: "text", text: '{"sheet":"Sheet1"}' }],
			isError: false,
		});

		const client = new McpStdioClient({ command: "npx", args: ["excel-mcp"] });
		const result = await client.callTool("list_sheets", {});

		expect(mockConnect).toHaveBeenCalledOnce();
		expect(result).toEqual({ sheet: "Sheet1" });
	});

	it("passes the correct command and args to StdioClientTransport", async () => {
		mockCallTool.mockResolvedValue({
			content: [{ type: "text", text: "{}" }],
		});

		const client = new McpStdioClient({
			command: "node",
			args: ["server.js", "--port", "3000"],
		});
		await client.callTool("ping", {});

		expect(mockTransportConstructor).toHaveBeenCalledWith({
			command: "node",
			args: ["server.js", "--port", "3000"],
		});
	});

	it("parses JSON text content from tool result", async () => {
		mockCallTool.mockResolvedValue({
			content: [{ type: "text", text: '{"rows":3}' }],
		});

		const client = new McpStdioClient({ command: "npx", args: ["mcp"] });
		const result = await client.callTool("count_rows", {});

		expect(result).toEqual({ rows: 3 });
	});

	it("returns raw string when tool result text is not valid JSON", async () => {
		mockCallTool.mockResolvedValue({
			content: [{ type: "text", text: "plain text result" }],
		});

		const client = new McpStdioClient({ command: "npx", args: ["mcp"] });
		const result = await client.callTool("get_text", {});

		expect(result).toBe("plain text result");
	});

	it("throws McpClientError when isError is true", async () => {
		mockCallTool.mockResolvedValue({
			content: [{ type: "text", text: "something went wrong" }],
			isError: true,
		});

		const client = new McpStdioClient({
			command: "npx",
			args: ["mcp"],
			name: "test-mcp",
		});

		await expect(client.callTool("fail", {})).rejects.toThrow(McpClientError);
		await expect(client.callTool("fail", {})).rejects.toMatchObject({
			code: "TOOL_ERROR",
		});
	});

	it("does not create a second connection on concurrent callTool() calls", async () => {
		mockCallTool.mockResolvedValue({
			content: [{ type: "text", text: "{}" }],
		});

		const client = new McpStdioClient({ command: "npx", args: ["mcp"] });

		await Promise.all([client.callTool("a", {}), client.callTool("b", {})]);

		expect(mockConnect).toHaveBeenCalledOnce();
	});

	it("close() calls client.close() and resets state", async () => {
		mockCallTool.mockResolvedValue({ content: [{ type: "text", text: "{}" }] });

		const client = new McpStdioClient({ command: "npx", args: ["mcp"] });
		await client.callTool("ping", {});
		await client.close();

		expect(mockClose).toHaveBeenCalledOnce();

		// After close, a new connection should be established on next callTool()
		await client.callTool("ping2", {});
		expect(mockConnect).toHaveBeenCalledTimes(2);
	});

	it("close() is safe to call before connecting", async () => {
		const client = new McpStdioClient({ command: "npx", args: ["mcp"] });
		await expect(client.close()).resolves.toBeUndefined();
		expect(mockClose).not.toHaveBeenCalled();
	});

	it("clears connectPromise on connect failure so a subsequent call retries", async () => {
		// First connect attempt fails.
		const connectError = new Error("spawn ENOENT");
		mockConnect.mockRejectedValueOnce(connectError);

		const client = new McpStdioClient({
			command: "npx",
			args: ["mcp"],
			name: "retry-test",
		});

		// First callTool() should reject with the connect error.
		await expect(client.callTool("ping", {})).rejects.toThrow("spawn ENOENT");

		// Subsequent connect must succeed; verify connectPromise was cleared.
		mockConnect.mockResolvedValue(undefined);
		mockCallTool.mockResolvedValue({
			content: [{ type: "text", text: '{"ok":true}' }],
		});

		const result = await client.callTool("ping", {});
		expect(result).toEqual({ ok: true });
		// connect() must have been called twice (once for the failed attempt, once for retry).
		expect(mockConnect).toHaveBeenCalledTimes(2);
	});
});

// ---------------------------------------------------------------------------
// McpClientError
// ---------------------------------------------------------------------------

describe("McpClientError", () => {
	it("stores code and has correct name", () => {
		const err = new McpClientError("oops", "CONNECTION_ERROR");
		expect(err.message).toBe("oops");
		expect(err.code).toBe("CONNECTION_ERROR");
		expect(err.name).toBe("McpClientError");
		expect(err instanceof Error).toBe(true);
	});
});
