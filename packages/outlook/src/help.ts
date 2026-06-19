/**
 * Why: runAxiCli needs a top-level help string and per-command help lookup; centralizing
 * them keeps usage text out of the command modules.
 * What: Exports TOP_LEVEL_HELP and a COMMAND_HELP record keyed by command name.
 * Test: Assert TOP_LEVEL_HELP contains "outlook-axi" and COMMAND_HELP["mail-list"] is set.
 */

export const TOP_LEVEL_HELP = `outlook-axi — AXI CLI for Outlook (via softeria/ms-365-mcp-server)

Usage: outlook-axi <command> [args] [flags]

Commands:
  mail-list     List mail messages
  mail-get      Get a single mail message
  mail-send     Send a mail message
  folders       List mail folders
  cal-list      List calendar events
  cal-view      List events in a date range
  cal-create    Create a calendar event
  contacts-list List contacts
  login         How to authenticate
  setup hooks   Install session-start hooks

Run \`outlook-axi <command> --help\` for per-command help.`;

export const COMMAND_HELP: Record<string, string> = {
  "mail-list": `outlook-axi mail-list [--limit N] [--folder NAME]

List mail messages. --folder restricts to a named folder.`,
  "mail-get": `outlook-axi mail-get <id>

Get a single mail message by id.`,
  "mail-send": `outlook-axi mail-send <to> <subject> <body>

Send a plain-text mail message to <to>.`,
  folders: `outlook-axi folders

List mail folders.`,
  "cal-list": `outlook-axi cal-list [--limit N]

List upcoming calendar events.`,
  "cal-view": `outlook-axi cal-view <start> <end>

List events between <start> and <end> (ISO 8601 datetimes).`,
  "cal-create": `outlook-axi cal-create <subject> <start> <end> [--attendees a@x,b@y]

Create a calendar event. --attendees is a comma-separated email list.`,
  "contacts-list": `outlook-axi contacts-list [--limit N]

List contacts.`,
  login: `outlook-axi login

Print guidance for authenticating with Microsoft 365 (device-code flow).`,
};
