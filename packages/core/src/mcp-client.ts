/**
 * Why: The @axi-office excel and outlook packages need to call tools on MCP servers
 * that run as local subprocess (e.g. negokaz/mcp-server-excel, softeria/mcp-office).
 * This client provides a thin, lifecycle-managed wrapper over the MCP SDK's
 * StdioClientTransport + Client so each package doesn't duplicate that setup.
 *
 * What: McpStdioClient spawns a server subprocess on first callTool(), reuses the
 * connection for the process lifetime, and closes cleanly on exit. It also exports a
 * standalone callHttpTool() for HTTP-based MCP servers (like lore-axi uses).
 *
 * Test: In unit tests, mock the StdioClientTransport constructor and Client class so
 * no real subprocess is spawned; assert that callTool() calls client.callTool with the
 * correct name and arguments, and that close() calls client.close().
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

// ---------------------------------------------------------------------------
// Typed result helpers
// ---------------------------------------------------------------------------

/** A single content block in an MCP tool result. */
export interface McpContentBlock {
  type: "text" | "image" | "resource";
  text?: string;
}

/** The raw result returned by the MCP SDK's client.callTool(). */
export interface McpRawResult {
  content: McpContentBlock[];
  isError?: boolean;
}

// ---------------------------------------------------------------------------
// McpStdioClient
// ---------------------------------------------------------------------------

/**
 * Options for constructing an McpStdioClient.
 */
export interface McpStdioClientOptions {
  /** Executable name or path for the MCP server (e.g. "npx" or "/usr/bin/node"). */
  command: string;
  /** Arguments to pass to the MCP server process. */
  args: string[];
  /** Display name used in error messages (e.g. "excel-mcp"). */
  name?: string;
  /** Version string sent in client info (defaults to "1.0.0"). */
  version?: string;
}

/**
 * Why: Provides lazy-connect lifecycle so the subprocess is only started when a
 * tool is first called, and is automatically torn down when the Node process exits.
 * What: Wraps StdioClientTransport + Client with connect()/callTool()/close().
 * Test: Mock StdioClientTransport and Client; construct McpStdioClient; call
 * callTool("list_sheets", {}); assert transport constructor was called with the
 * correct command/args and client.callTool was invoked with the right params.
 */
export class McpStdioClient {
  private readonly opts: Required<McpStdioClientOptions>;
  private client: Client | undefined;
  private transport: StdioClientTransport | undefined;
  private connectPromise: Promise<void> | undefined;

  constructor(options: McpStdioClientOptions) {
    this.opts = {
      name: options.name ?? options.command,
      version: options.version ?? "1.0.0",
      command: options.command,
      args: options.args,
    };
  }

  /**
   * Why: Ensures a single connection is established even when multiple callTool()
   * calls race during startup.
   * What: Lazily initializes the StdioClientTransport and Client, then calls
   * client.connect(). Stores the in-flight promise to prevent duplicate connections.
   * Test: Call connect() twice concurrently; assert the Client constructor is only
   * called once.
   */
  async connect(): Promise<void> {
    if (this.client) return;
    if (this.connectPromise) {
      await this.connectPromise;
      return;
    }

    this.connectPromise = (async () => {
      const transport = new StdioClientTransport({
        command: this.opts.command,
        args: this.opts.args,
      });

      const client = new Client(
        { name: this.opts.name, version: this.opts.version },
        { capabilities: {} }
      );

      await client.connect(transport);

      this.transport = transport;
      this.client = client;

      // Ensure the subprocess is torn down when Node exits.
      process.once("exit", () => void this.close());
    })();

    await this.connectPromise;
  }

  /**
   * Why: Provides the primary API for consumers — call an MCP tool and get a typed
   * result without dealing with transport or protocol details.
   * What: Lazily connects on first call, then delegates to client.callTool(). Parses
   * the first text content block as JSON when possible; returns the raw string
   * otherwise. Throws McpClientError on tool errors.
   * Test: Mock client.callTool to return { content: [{ type: "text", text: '{"a":1}' }] };
   * assert callTool returns { a: 1 }.
   */
  async callTool(name: string, args: Record<string, unknown> = {}): Promise<unknown> {
    await this.connect();

    if (!this.client) {
      throw new McpClientError(`${this.opts.name}: client is not connected`, "NOT_CONNECTED");
    }

    const raw = (await this.client.callTool({
      name,
      arguments: args,
    })) as McpRawResult;

    if (raw.isError) {
      const errText = raw.content?.[0]?.text ?? "Unknown tool error";
      throw new McpClientError(`${this.opts.name}: ${errText}`, "TOOL_ERROR");
    }

    const text = raw.content?.[0]?.text;
    if (text === undefined || text === null) return {};

    try {
      return JSON.parse(text) as unknown;
    } catch {
      return text;
    }
  }

  /**
   * Why: Prevents zombie subprocesses by explicitly closing the transport.
   * What: Calls client.close() and clears internal state. Safe to call multiple times.
   * Test: Call close() twice; assert client.close() is only called once.
   */
  async close(): Promise<void> {
    if (!this.client) return;
    try {
      await this.client.close();
    } catch {
      // Ignore close errors during teardown.
    }
    this.client = undefined;
    this.transport = undefined;
    this.connectPromise = undefined;
  }
}

// ---------------------------------------------------------------------------
// Structured error
// ---------------------------------------------------------------------------

/** Error thrown by McpStdioClient when the transport or tool call fails. */
export class McpClientError extends Error {
  readonly code: string;

  constructor(message: string, code: string) {
    super(message);
    this.name = "McpClientError";
    this.code = code;
  }
}

// ---------------------------------------------------------------------------
// HTTP MCP helper (for servers that use HTTP transport like lore-axi)
// ---------------------------------------------------------------------------

/**
 * Why: Some MCP servers (like lore) expose an HTTP endpoint rather than stdio.
 * This helper provides a single-round-trip HTTP call without requiring a persistent
 * client instance.
 * What: Sends a JSON-RPC 2.0 POST to url with method "tools/call", parses the
 * response content, and returns the parsed result.
 * Test: Mock fetch to return a valid JSON-RPC response; assert the parsed text is
 * returned as a plain object.
 */
export async function callHttpTool(
  url: string,
  name: string,
  args: Record<string, unknown> = {}
): Promise<unknown> {
  const body = JSON.stringify({
    jsonrpc: "2.0",
    method: "tools/call",
    params: { name, arguments: args },
    id: 1,
  });

  let response: Response;
  try {
    response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body,
    });
  } catch (err) {
    throw new McpClientError(
      `Failed to connect to MCP server at ${url}: ${String(err)}`,
      "CONNECTION_ERROR"
    );
  }

  if (!response.ok) {
    throw new McpClientError(
      `MCP server returned ${response.status} ${response.statusText}`,
      "HTTP_ERROR"
    );
  }

  let data: unknown;
  try {
    data = await response.json();
  } catch {
    throw new McpClientError("MCP server returned non-JSON response", "PARSE_ERROR");
  }

  const rpc = data as {
    error?: { message: string; code: number };
    result?: McpRawResult;
  };

  if (rpc.error) {
    throw new McpClientError(
      `MCP error: ${rpc.error.message} (code ${rpc.error.code})`,
      "MCP_ERROR"
    );
  }

  const result = rpc.result;
  if (!result) {
    throw new McpClientError("MCP response missing result", "MCP_ERROR");
  }

  if (result.isError) {
    const errText = result.content?.[0]?.text ?? "Unknown error";
    throw new McpClientError(errText, "MCP_TOOL_ERROR");
  }

  const text = result.content?.[0]?.text;
  if (text === undefined || text === null) return {};

  try {
    return JSON.parse(text) as unknown;
  } catch {
    return text;
  }
}
