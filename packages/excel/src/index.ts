/**
 * Why: Exposes the excel command handlers and client factory for programmatic use and
 * for the bin entry point to import from a single module path.
 * What: Re-exports every command handler, the client factory, and help strings.
 * Test: Import { sheetsCommand, getClient, TOP_LEVEL_HELP } from "@axi-office/excel"
 * and assert they are defined.
 */
export { getClient } from "./client.js";
export { TOP_LEVEL_HELP, COMMAND_HELP } from "./help.js";
export { homeCommand } from "./home.js";
export { sheetsCommand } from "./commands/sheets.js";
export { readCommand } from "./commands/read.js";
export { writeCommand } from "./commands/write.js";
export { createTableCommand } from "./commands/create-table.js";
export { copySheetCommand } from "./commands/copy-sheet.js";
export { formatRangeCommand } from "./commands/format-range.js";
