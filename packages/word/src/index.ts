/**
 * Why: Exposes the word command handlers and helpers for programmatic use and for the bin
 * entry to import from a single module path.
 * What: Re-exports every command handler, the docx builder, the markdown parser, and help.
 * Test: Import { createCommand, buildDocxBuffer, parseMarkdown } from "@axi-office/word"
 * and assert they are defined.
 */
export { TOP_LEVEL_HELP, COMMAND_HELP } from "./help.js";
export { homeCommand } from "./home.js";
export { buildDocxBuffer } from "./docx-build.js";
export type { DocSpec, DocSection, InlineRun } from "./docx-build.js";
export { parseMarkdown, parseInline } from "./markdown.js";
export { createCommand } from "./commands/create.js";
export { fromMarkdownCommand } from "./commands/from-markdown.js";
export { readCommand } from "./commands/read.js";
export { infoCommand } from "./commands/info.js";
export { patchCommand } from "./commands/patch.js";
