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
	/**
	 * Milliseconds to wait for the SDK connect (subprocess spawn + handshake)
	 * before rejecting with a CONNECT_TIMEOUT error. Defaults to 30000. Set to 0
	 * (or a negative value) to disable the timeout entirely.
	 */
	connectTimeoutMs?: number;
	/**
	 * How to handle stderr of the child MCP process. Mirrors Node's child_process
	 * spawn stdio options (e.g. "inherit", "pipe", "ignore").
	 *
	 * Defaults to "inherit" (child stderr appears on the parent's terminal).
	 * Set to "ignore" to suppress server-side tracebacks that some MCP servers
	 * (e.g. haris-musa/excel-mcp-server) print on clean shutdown due to an
	 * upstream bug where `print("Service stopped.")` raises ValueError when the
	 * stdio pipe is already closed.
	 * Set to "pipe" to capture stderr programmatically via the transport's
	 * .stderr stream.
	 */
	stderr?: "inherit" | "pipe" | "ignore" | "overlapped";
}

/**
 * Why: Provides lazy-connect lifecycle so the subprocess is only started when a
 * tool is first called, and is automatically torn down when the Node process exits.
 * What: Wraps StdioClientTransport + Client with connect()/callTool()/close().
 * Test: Mock StdioClientTransport and Client; construct McpStdioClient; call
 * callTool("list_sheets", {}); assert transport constructor was called with the
 * correct command/args and client.callTool was invoked with the right params.
 */
// ---------------------------------------------------------------------------
// Module-level process-exit registry — ensures we register signal handlers
// exactly once across all McpStdioClient instances.
// ---------------------------------------------------------------------------

/** All live client instances, used by the module-level signal handlers. */
const _liveClients = new Set<McpStdioClient>();
let _signalHandlersRegistered = false;

/**
 * Why: The live-clients registry is module-private (not part of the public API),
 * but tests must assert that connect()/close() add and remove instances so that
 * signal-handler teardown only touches still-open clients.
 * What: Returns the internal live-clients Set for inspection. Test-only; not
 * re-exported from the package index.
 * Test: Used by the test suite to assert registry membership after connect/close.
 */
export function _getLiveClientsForTest(): ReadonlySet<McpStdioClient> {
	return _liveClients;
}

/**
 * Why: Bins need to await clean teardown of all spawned MCP child processes
 * before calling process.exit(), so that the child flushes stderr and exits
 * cleanly before Node closes the stdio fds. Without this, the Python MCP
 * server prints "ValueError: I/O operation on closed file" / "Service stopped."
 * tracebacks to the terminal after every command.
 * What: Calls close() on every client currently in the live-clients registry,
 * swallowing individual errors via Promise.allSettled so one stuck client does
 * not block others from being torn down.
 * Test: Register two McpStdioClient instances; call closeAllLiveClients(); assert
 * both clients' close() was called and the registry is empty afterward.
 */
export async function closeAllLiveClients(): Promise<void> {
	await Promise.allSettled(Array.from(_liveClients).map((c) => c.close()));
}

/**
 * Whether the module-level signal handlers call process.exit() after closing
 * clients. Defaults to true (CLI behavior). Toggled via setAutoExitOnSignal().
 */
let autoExitOnSignal = true;

/** POSIX exit codes for terminated-by-signal processes: 128 + signal number. */
const _EXIT_CODE_SIGTERM = 143; // 128 + 15
const _EXIT_CODE_SIGINT = 130; // 128 + 2

/** Guard so the "_process unavailable" warning is emitted at most once. */
let _warnedProcessUnavailable = false;

/**
 * Why: The synchronous "exit" backstop can only SIGKILL a stuck subprocess when
 * it holds the ChildProcess reference (transport._process). If a future SDK
 * version renames or removes that private field, auto-kill silently degrades to
 * relying on transport.close() alone — we surface that once so operators know.
 * What: Emits a single stderr warning the first time _process is unavailable.
 * Test: Mock a transport without _process, connect, and assert console.warn was
 * called once with the expected message; connect a second client and assert it
 * is not called again.
 */
function _warnProcessUnavailableOnce(): void {
	if (_warnedProcessUnavailable) return;
	_warnedProcessUnavailable = true;
	console.warn(
		"[axi-office/core] subprocess auto-kill unavailable: transport._process not accessible"
	);
}

/**
 * Why: When this client is embedded in a long-lived host process (a server,
 * test runner, or another CLI that owns its own lifecycle), having the MCP
 * client forcibly call process.exit() on SIGINT/SIGTERM would steal control of
 * shutdown from the host. This toggle lets the host opt out of the auto-exit
 * behavior while still benefiting from automatic child-process cleanup.
 * What: Sets whether the SIGINT/SIGTERM handlers call process.exit() after
 * tearing down live clients. Defaults to true (CLI behavior: the MCP tool is
 * the process, so exiting with a POSIX signal code is correct). Set to false
 * for library embedding where the host manages the process lifecycle — the
 * handlers still kill child processes and clear the registry, but never exit.
 * Test: Call setAutoExitOnSignal(false); emit "SIGINT" on a mocked process and
 * assert process.exit was NOT called but child processes were killed and the
 * registry was emptied.
 */
export function setAutoExitOnSignal(enabled: boolean): void {
	autoExitOnSignal = enabled;
}

function _ensureSignalHandlers(): void {
	if (_signalHandlersRegistered) return;
	_signalHandlersRegistered = true;

	// Graceful shutdown on SIGINT / SIGTERM: await close() then (optionally) exit
	// with the POSIX 128+signal code. When autoExitOnSignal is false (library
	// embedding) we still tear down children but leave process lifecycle to the host.
	async function gracefulShutdown(signal: "SIGINT" | "SIGTERM"): Promise<void> {
		const closes = Array.from(_liveClients).map((c) => c.close());
		await Promise.allSettled(closes);
		if (!autoExitOnSignal) return;
		process.exit(signal === "SIGTERM" ? _EXIT_CODE_SIGTERM : _EXIT_CODE_SIGINT);
	}

	process.on("SIGINT", () => void gracefulShutdown("SIGINT"));
	process.on("SIGTERM", () => void gracefulShutdown("SIGTERM"));

	// Synchronous backstop: if the process exits for any reason (unhandled
	// exception, process.exit() called elsewhere, etc.) synchronously kill
	// every child that is still alive. async close() is not possible here.
	process.on("exit", () => {
		for (const client of _liveClients) {
			client._syncKillChild();
		}
	});
}

export class McpStdioClient {
	private readonly opts: Required<McpStdioClientOptions>;
	private client: Client | undefined;
	private transport: StdioClientTransport | undefined;
	private connectPromise: Promise<void> | undefined;
	/** Reference kept so the "exit" handler can synchronously kill the child. */
	private _childProcess: import("node:child_process").ChildProcess | undefined;

	constructor(options: McpStdioClientOptions) {
		this.opts = {
			name: options.name ?? options.command,
			version: options.version ?? "1.0.0",
			command: options.command,
			args: options.args,
			connectTimeoutMs: options.connectTimeoutMs ?? 30000,
			stderr: options.stderr ?? "inherit",
		};
	}

	/**
	 * Why: Called synchronously from the process "exit" handler as a best-effort
	 * backstop to kill the spawned child when async close() cannot be awaited.
	 * What: Sends SIGKILL to the child process if it is still alive.
	 * Test: Not tested directly — covered by the lifecycle integration tests.
	 */
	_syncKillChild(): void {
		try {
			this._childProcess?.kill("SIGKILL");
		} catch {
			// Ignore errors during forced kill.
		}
	}

	/**
	 * Why: Ensures a single connection is established even when multiple callTool()
	 * calls race during startup.
	 * What: Lazily initializes the StdioClientTransport and Client, then calls
	 * client.connect(). Stores the in-flight promise to prevent duplicate connections.
	 * If connect fails, clears connectPromise so the next call retries with a fresh
	 * attempt rather than permanently awaiting a rejected promise.
	 * Test: Call connect() twice concurrently; assert the Client constructor is only
	 * called once. Also: make connect fail once, assert callTool() rejects, then make
	 * it succeed and assert the next callTool() works (retry path). Also: with a
	 * transport whose connect never resolves, connect({ connectTimeoutMs: 100 })
	 * rejects with a CONNECT_TIMEOUT error and clears connectPromise so retry works.
	 *
	 * @param overrides Optional per-call overrides (e.g. connectTimeoutMs) that take
	 * precedence over the constructor options for this connect attempt only.
	 */
	async connect(overrides?: { connectTimeoutMs?: number }): Promise<void> {
		if (this.client) return;
		if (this.connectPromise) {
			await this.connectPromise;
			return;
		}

		const connectTimeoutMs = overrides?.connectTimeoutMs ?? this.opts.connectTimeoutMs;

		this.connectPromise = (async () => {
			try {
				const transport = new StdioClientTransport({
					command: this.opts.command,
					args: this.opts.args,
					stderr: this.opts.stderr,
				});

				const client = new Client(
					{ name: this.opts.name, version: this.opts.version },
					{ capabilities: {} }
				);

				// Wrap the SDK connect with a timeout so a server that spawns but never
				// completes the handshake cannot hang the caller indefinitely.
				const doConnect = () => client.connect(transport);
				if (connectTimeoutMs > 0) {
					let timer: ReturnType<typeof setTimeout> | undefined;
					const timeout = new Promise<never>((_, reject) => {
						timer = setTimeout(() => {
							reject(
								new McpClientError(
									`${this.opts.name}: MCP connect timed out after ${connectTimeoutMs}ms`,
									"CONNECT_TIMEOUT"
								)
							);
						}, connectTimeoutMs);
					});
					try {
						await Promise.race([doConnect(), timeout]);
					} finally {
						if (timer) clearTimeout(timer);
					}
				} else {
					await doConnect();
				}

				// Capture the child process reference for the synchronous exit backstop.
				// StdioClientTransport stores the spawned process as ._process (private,
				// unofficial but stable in MCP SDK ≥ 1.x); guard defensively.
				// Note: the public .pid getter exists, but we need the full ChildProcess
				// to call .kill() synchronously in the exit handler.
				const maybeChild = (
					transport as unknown as {
						_process?: import("node:child_process").ChildProcess;
					}
				)._process;
				if (maybeChild) {
					this._childProcess = maybeChild;
				} else {
					// Observability: without the child reference the synchronous exit
					// backstop cannot SIGKILL a stuck subprocess. Warn once to stderr.
					_warnProcessUnavailableOnce();
				}

				this.transport = transport;
				this.client = client;

				// Register this client for module-level signal/exit handling.
				_liveClients.add(this);
				_ensureSignalHandlers();
			} catch (err) {
				// Clear the promise so a subsequent connect() call can retry rather
				// than permanently awaiting this rejected promise.
				this.connectPromise = undefined;
				throw err;
			}
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
	 * Why: Prevents zombie subprocesses by explicitly closing the transport. If
	 * close() is called while a connect is still in-flight, we must let that
	 * connect settle first — otherwise we would return from close() with a child
	 * that is still starting up (a zombie), because this.client is not yet set.
	 * What: Awaits any in-flight connectPromise (swallowing its error), then calls
	 * client.close() and clears internal state. Safe to call multiple times and
	 * safe to call before connecting.
	 * Test: Call close() twice; assert client.close() is only called once. Also:
	 * start connect() with a transport that resolves slowly, immediately call
	 * close(), and assert close() resolves only after the connect settles and the
	 * child is torn down (client removed from the registry).
	 */
	async close(): Promise<void> {
		// If a connect is in-flight, let it settle before tearing down so we never
		// return from close() while the child process is still starting up.
		if (this.connectPromise) {
			await this.connectPromise.catch(() => {});
		}
		if (!this.client) return;
		try {
			await this.client.close();
		} catch {
			// Ignore close errors during teardown.
		}
		this.client = undefined;
		this.transport = undefined;
		this.connectPromise = undefined;
		this._childProcess = undefined;
		_liveClients.delete(this);
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
 *
 * @internal
 * Note: performs no SSRF guard — the caller supplies the URL. Never pass
 * untrusted user input as the `url` argument.
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
