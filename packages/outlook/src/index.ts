/**
 * Why: Exposes the outlook command handlers and client factory for programmatic use and for
 * the bin entry to import from a single module path.
 * What: Re-exports every command handler, the client factory, and help strings.
 * Test: Import { mailListCommand, getClient, TOP_LEVEL_HELP } from "@axi-office/outlook"
 * and assert they are defined.
 */
export { getClient } from "./client.js";
export { TOP_LEVEL_HELP, COMMAND_HELP } from "./help.js";
export { homeCommand } from "./home.js";
export {
	mailListCommand,
	mailGetCommand,
	mailSendCommand,
	foldersCommand,
} from "./commands/mail.js";
export {
	calListCommand,
	calViewCommand,
	calCreateCommand,
} from "./commands/calendar.js";
export { contactsListCommand } from "./commands/contacts.js";
export { loginCommand } from "./commands/login.js";
