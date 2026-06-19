/**
 * Why: Contacts are the third Outlook surface; this maps contacts-list to the
 * ms-365-mcp-server list-outlook-contacts tool.
 * What: contactsListCommand forwards an optional --limit as $top to list-outlook-contacts.
 * Test: Mock the client, call contactsListCommand(["--limit","10"]); assert callTool was
 * invoked with "list-outlook-contacts" and $top === 10.
 */
import { parseFlags, parseLimit } from "@axi-office/core";
import { getClient } from "../client.js";

export async function contactsListCommand(args: string[]): Promise<unknown> {
  const { flags } = parseFlags(args);
  const toolArgs: Record<string, unknown> = {};
  if (flags.limit !== undefined) {
    toolArgs.$top = parseLimit(flags.limit, 25, 100);
  }
  return getClient().callTool("list-outlook-contacts", toolArgs);
}
