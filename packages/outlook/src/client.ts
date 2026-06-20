/**
 * Why: All outlook commands talk to the same ms-365-mcp-server subprocess; centralizing
 * client construction avoids spawning duplicate servers and lets tests mock one factory.
 * What: Lazily constructs and memoizes one McpStdioClient running
 * `npx -y @softeria/ms-365-mcp-server --preset outlook`. The subprocess inherits the
 * parent process env (StdioClientTransport default), so MS365_MCP_* auth variables set in
 * the shell are passed through automatically.
 * Test: Mock McpStdioClient, call getClient() twice, assert the constructor ran once with
 * the outlook preset args.
 */
import { McpStdioClient } from "@axi-office/core";

let _client: McpStdioClient | undefined;

export function getClient(): McpStdioClient {
	if (!_client) {
		_client = new McpStdioClient({
			command: "npx",
			args: ["-y", "@softeria/ms-365-mcp-server@0.125.1", "--preset", "outlook"],
			// Pinned for supply-chain safety; audit upstream changelog before bumping.
			name: "ms-365-mcp",
		});
	}
	return _client;
}
