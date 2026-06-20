# @axi-office/outlook

AXI CLI for Outlook — a token-efficient wrapper around
[softeria/ms-365-mcp-server](https://www.npmjs.com/package/@softeria/ms-365-mcp-server),
scoped to the `outlook` preset (mail, calendar, contacts).

## Disclaimer

Unofficial project — not affiliated with or endorsed by Microsoft Corporation.
This package wraps the following open-source MCP servers:
- [softeria/ms-365-mcp-server](https://www.npmjs.com/package/@softeria/ms-365-mcp-server) (for Outlook)

All product names and trademarks are property of their respective owners.

## Install

```bash
npm install -g @axi-office/outlook
```

The CLI auto-fetches the pinned server (`@softeria/ms-365-mcp-server@0.125.1`) via npx on first run (network required once). You may pre-install a pinned global to avoid auto-fetch: `npm install -g @softeria/ms-365-mcp-server@0.125.1`

## Authenticate

```bash
npx -y @softeria/ms-365-mcp-server --login
```

Follow the device-code URL/code to sign in. Tokens are cached by the MCP server.
`outlook-axi login` prints this guidance. Optional environment variables (inherited by
the MCP subprocess when set in your shell):

| Variable | Purpose |
|----------|---------|
| `MS365_MCP_CLIENT_ID` | Custom Entra app client id |
| `MS365_MCP_TENANT_ID` | Tenant id |
| `MS365_MCP_OAUTH_TOKEN` | Pre-issued OAuth token |

## Custom Entra app registration

To use your own Azure Entra (formerly Azure AD) app instead of the default:

1. Create a **public client** app registration in the [Azure Portal](https://portal.azure.com) (Entra ID > App registrations > New registration).
2. Under **Authentication**, enable **"Allow public client flows"** (required for device-code auth).
3. Add the redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`
4. Under **API permissions**, grant only the least-privilege delegated scopes your use case requires:
   - `Mail.Read` — for reading mail
   - `Mail.Send` — for sending mail
   - `Calendars.Read` — for reading calendar events
   - `Calendars.ReadWrite` — for creating/updating events
   - `Contacts.Read` — for reading contacts
5. If your tenant requires it, have an admin grant consent for the permissions.
6. Set `MS365_MCP_CLIENT_ID` and `MS365_MCP_TENANT_ID` in your shell before running `outlook-axi`.

## Token cache & revocation

The ms-365-mcp-server caches OAuth tokens on disk (location is managed by the server).
To revoke and clear the cached token, run:

```bash
npx -y @softeria/ms-365-mcp-server@0.125.1 --logout
```

You can also revoke access from the [Microsoft My Apps portal](https://myapps.microsoft.com) or via Entra admin consent management.

## Bearer token security

`MS365_MCP_OAUTH_TOKEN` is a bearer token with a typical lifetime of 60–90 minutes.
**Treat it as a secret:**

- Do not persist it in dotfiles (`.bashrc`, `.zshrc`, `.env` files checked into source control).
- Do not commit it to any repository.
- If it is accidentally exposed, revoke it immediately via the Microsoft portal and rotate your credentials.

## Usage

```bash
outlook-axi <command> [args] [flags]
```

| Command | Description |
|---------|-------------|
| `mail-list [--limit N] [--folder NAME]` | List mail messages |
| `mail-get <id>` | Get a single mail message |
| `mail-send <to> <subject> <body>` | Send a plain-text message |
| `folders` | List mail folders |
| `cal-list [--limit N]` | List calendar events |
| `cal-view <start> <end>` | List events in an ISO datetime range |
| `cal-create <subject> <start> <end> [--attendees a@x,b@y]` | Create an event |
| `contacts-list [--limit N]` | List contacts |
| `login` | Print authentication guidance |
| `setup hooks` | Install AXI session-start hooks |

### Examples

```bash
outlook-axi mail-list --limit 10 --folder Inbox
outlook-axi mail-get AAMkAGI...
outlook-axi mail-send dev@example.com "Hello" "Body text"
outlook-axi cal-view 2026-01-01T00:00:00Z 2026-01-31T00:00:00Z
outlook-axi cal-create "Standup" 2026-02-01T09:00:00 2026-02-01T09:30:00 --attendees a@x.com,b@y.com
outlook-axi contacts-list --limit 25
```

Output is rendered as [TOON](https://github.com/toon-format/toon).

## License

MIT
