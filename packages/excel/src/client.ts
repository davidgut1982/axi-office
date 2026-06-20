/**
 * Why: The excel commands all talk to the same haris-musa/excel-mcp-server subprocess;
 * centralizing client construction avoids spawning duplicate servers and lets tests
 * mock a single factory.
 * What: Lazily constructs and memoizes one McpStdioClient pointed at
 * `uvx excel-mcp-server@0.1.8 stdio`.
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
			args: ["excel-mcp-server@0.1.8", "stdio"],
			name: "excel-mcp",
		});
	}
	return _client;
}
