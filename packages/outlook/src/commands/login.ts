/**
 * Why: ms-365-mcp-server uses an interactive device-code OAuth flow that cannot run inside
 * a single non-interactive tool call; the safe behavior is to guide the user rather than
 * attempt an automated login.
 * What: loginCommand returns guidance describing how to authenticate and which env vars to
 * set, without spawning the MCP server.
 * Test: Call loginCommand([]) and assert the result contains a `steps` array mentioning
 * `--login`.
 */

export async function loginCommand(_args: string[]): Promise<unknown> {
	return {
		auth: "Microsoft 365 device-code flow (handled by ms-365-mcp-server)",
		steps: [
			"Run: npx -y @softeria/ms-365-mcp-server --login",
			"Follow the printed device-code URL and code to sign in",
			"Tokens are cached by the MCP server for subsequent outlook-axi calls",
		],
		env: [
			"MS365_MCP_CLIENT_ID    — optional custom Entra app client id",
			"MS365_MCP_TENANT_ID    — optional tenant id",
			"MS365_MCP_OAUTH_TOKEN  — optional pre-issued OAuth token",
		],
		note: "These env vars, when set in your shell, are inherited by the MCP subprocess.",
	};
}
