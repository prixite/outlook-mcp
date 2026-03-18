# outlook-mcp

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io) server that lets AI agents manage **Microsoft Outlook on macOS** — read and send email, manage calendar events, and more.

## Requirements

- macOS with Microsoft Outlook installed
- Node.js ≥ 18
- Outlook must be running and accessible via AppleScript

## Installation

```bash
npm install
npm run build
```

## MCP Client Configuration

Add to your MCP client config (e.g. Claude Desktop `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/absolute/path/to/outlook-mcp/dist/index.js"]
    }
  }
}
```

Or run in dev mode (no build needed):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["tsx", "/absolute/path/to/outlook-mcp/src/index.ts"]
    }
  }
}
```

## Available Tools

### Email

| Tool | Description |
|------|-------------|
| `outlook_list_messages` | List recent messages from a folder (inbox, sent, drafts, deleted, or custom) |
| `outlook_read_message` | Read the full content of a message by ID |
| `outlook_search_messages` | Search inbox by keyword (subject, sender, preview) |
| `outlook_send_email` | Compose and send an email |
| `outlook_reply` | Reply to a message (supports reply-all) |
| `outlook_forward` | Forward a message to one or more recipients |
| `outlook_delete_message` | Move a message to Deleted Items |
| `outlook_mark_read` | Mark a message as read or unread |
| `outlook_list_folders` | List all mail folders with unread counts |
| `outlook_list_accounts` | List all configured email accounts (Exchange, IMAP, POP) |

### Calendar

| Tool | Description |
|------|-------------|
| `outlook_list_events` | List upcoming calendar events |
| `outlook_get_event` | Get full details of a calendar event by ID |
| `outlook_create_event` | Create a new calendar event |

## Development

```bash
npm run dev          # Run with tsx (no build step)
npm run build        # Compile TypeScript
npm run lint         # Check for lint errors
npm run lint:fix     # Auto-fix lint errors
npm run format       # Format source files
npm run format:check # Check formatting without writing
```

## How It Works

Each tool call generates an AppleScript, executes it via `osascript`, and parses the output (tab-separated values) back into structured JSON. No external APIs or OAuth tokens are required — it communicates directly with the Outlook macOS app.

## License

MIT
