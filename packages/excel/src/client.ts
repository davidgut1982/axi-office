/**
 * Why: The excel commands all talk to the same negokaz/excel-mcp-server subprocess;
 * centralizing client construction avoids spawning duplicate servers and lets tests
 * mock a single factory.
 * What: Lazily constructs and memoizes one McpStdioClient pointed at
 * `npx -y @negokaz/excel-mcp-server`.
 * Test: Mock McpStdioClient, call getClient() twice, assert the constructor ran once
 * and the same instance is returned.
 */
import { McpStdioClient } from "@axi-office/core";

let _client: McpStdioClient | undefined;

export function getClient(): McpStdioClient {
	if (!_client) {
		_client = new McpStdioClient({
			command: "npx",
			args: ["-y", "@negokaz/excel-mcp-server"],
			name: "excel-mcp",
		});
	}
	return _client;
}
