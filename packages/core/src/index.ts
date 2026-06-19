/**
 * Why: Provides a single entry point so consumers can import common plumbing from
 * "@axi-office/core" without knowing internal module paths.
 * What: Re-exports the AXI SDK primitives plus the core helpers (toon, args, mcp-client, hooks).
 * Test: Import { runAxiCli, AxiError, parseFlags, field, McpStdioClient, setupHooksCommand }
 * from "@axi-office/core" in a downstream package and assert they are functions/classes.
 */

// axi-sdk-js re-exports
export { runAxiCli, installSessionStartHooks } from "axi-sdk-js";
export { AxiError } from "axi-sdk-js";

// TOON field-schema helpers
export {
  field,
  pluck,
  joinArray,
  relativeTime,
  boolYesNo,
  mapEnum,
  lower,
  custom,
  extract,
  renderList,
  renderDetail,
  renderHelp,
  renderOutput,
  formatRelativeTime,
} from "./toon.js";

export type {
  FieldDef,
  PluckDef,
  JoinArrayDef,
  RelativeTimeDef,
  BoolYesNoDef,
  MapEnumDef,
  LowerDef,
  CustomDef,
  SchemaDef,
} from "./toon.js";

// Flag/arg parsing helpers
export {
  parseFlags,
  parseLimit,
  parseKeyValuePairs,
  coerceValue,
  collapseWhitespace,
  truncateLine,
} from "./args.js";

export type { ParsedArgs } from "./args.js";

// MCP STDIO client
export { McpStdioClient, McpClientError, callHttpTool } from "./mcp-client.js";

export type {
  McpStdioClientOptions,
  McpContentBlock,
  McpRawResult,
} from "./mcp-client.js";

// Session hooks helper
export { setupHooksCommand } from "./hooks.js";
