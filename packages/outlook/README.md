# @axi-office/outlook

AXI CLI for Outlook — a token-efficient wrapper around
[softeria/ms-365-mcp-server](https://www.npmjs.com/package/@softeria/ms-365-mcp-server),
scoped to the `outlook` preset (mail, calendar, contacts).

## Install

```bash
npm install -g @axi-office/outlook
```

The MCP server is fetched on demand via `npx`.

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
