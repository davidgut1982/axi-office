/**
 * Why: The excel commands all talk to the same haris-musa/excel-mcp-server subprocess;
 * centralizing client construction avoids spawning duplicate servers and lets tests
 * mock a single factory.
 * What: Lazily constructs and memoizes one McpStdioClient pointed at
 * `uvx excel-mcp-server@0.1.8 stdio`. stderr is set to "ignore" to suppress the
 * "Service stopped." / ValueError traceback that excel-mcp-server@0.1.8 always prints
 * on clean shutdown due to an upstream bug where the Python finally block tries to
 * print("Service stopped.") to a stdout pipe that is already closed. This is a
 * cosmetic suppression — the traceback carries no actionable information.
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
			// Suppress the "Service stopped." / ValueError traceback that
			// excel-mcp-server@0.1.8 always emits on clean shutdown because its
			// Python finally block writes to a closed stdout pipe.
			stderr: "ignore",
		});
	}
	return _client;
}
