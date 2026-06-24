/**
 * Why: Exposes the ppt command handlers and client factory for programmatic use and
 * for the bin entry point to import from a single module path.
 * What: Re-exports every command handler, the client factory, session helpers, and help strings.
 * Test: Import { createCommand, getClient, TOP_LEVEL_HELP } from "@axi-office/ppt"
 * and assert they are defined.
 */
export { getClient } from "./client.js";
export { withOpenSave, withOpenReadonly, withCreateSave } from "./session.js";
export { TOP_LEVEL_HELP, COMMAND_HELP } from "./help.js";
export { homeCommand } from "./home.js";
export { createCommand } from "./commands/create.js";
export { infoCommand } from "./commands/info.js";
export { setPropsCommand } from "./commands/set-props.js";
export { fromTemplateCommand } from "./commands/from-template.js";
export { addSlideCommand } from "./commands/add-slide.js";
export { slideInfoCommand } from "./commands/slide-info.js";
export { readCommand } from "./commands/read.js";
export { readSlideCommand } from "./commands/read-slide.js";
export { setPlaceholderCommand } from "./commands/set-placeholder.js";
export { bulletsCommand } from "./commands/bullets.js";
export { addTextCommand } from "./commands/add-text.js";
export { addImageCommand } from "./commands/add-image.js";
export { addTableCommand } from "./commands/add-table.js";
export { addShapeCommand } from "./commands/add-shape.js";
export { addChartCommand } from "./commands/add-chart.js";
export { templatesCommand } from "./commands/templates.js";
export { autoGenerateCommand } from "./commands/auto-generate.js";
