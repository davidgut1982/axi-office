/**
 * Why: All ppt commands talk to the same GongRzhe/Office-PowerPoint-MCP-Server subprocess;
 * centralizing client construction avoids spawning duplicate servers and lets tests
 * mock a single factory.
 * What: Lazily constructs and memoizes one McpStdioClient pointed at
 * `uvx --from office-powerpoint-mcp-server==2.0.7 ppt_mcp_server`. stderr is set to
 * "ignore" to suppress any traceback noise emitted on clean shutdown.
 * Test: Mock McpStdioClient, call getClient() twice, assert the constructor ran once
 * and the same instance is returned.
 */
import { McpStdioClient } from "@axi-office/core";

let _client: McpStdioClient | undefined;

export function getClient(): McpStdioClient {
	if (!_client) {
		_client = new McpStdioClient({
			command: "uvx",
			// Pinned for supply-chain safety; audit upstream changelog before bumping.
			args: ["--from", "office-powerpoint-mcp-server==2.0.7", "ppt_mcp_server"],
			name: "ppt-mcp",
			// Suppress any cleanup tracebacks from the Python server on shutdown.
			stderr: "ignore",
		});
	}
	return _client;
}
