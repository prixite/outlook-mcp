# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run build    # Compile TypeScript to dist/
npm run dev      # Run server in development mode (tsx, no build needed)
npm start        # Run compiled server: node dist/index.js
```

There is no test or lint setup. Before building, ensure TypeScript compiles cleanly with `npm run build`.

## Architecture

This is an MCP (Model Context Protocol) server that bridges AI agents to Microsoft Outlook on macOS via AppleScript.

**Data flow:**
1. MCP client sends a tool call over stdio
2. Tool handler validates input with Zod
3. Handler generates an AppleScript, writes it to a temp file, and executes it via `osascript`
4. Output (TSV-formatted) is parsed and returned as JSON to the MCP client

**Key modules:**
- `src/index.ts` — Server entry point; registers tools and starts stdio transport
- `src/applescript.ts` — Core utilities: `runAppleScript()`, `asString()` (safe escaping), `parseTSV()`, `flattenForTSV()`
- `src/tools/email-tools.ts` — 9 email tools (list, read, search, send, reply, forward, delete, mark-read, list-folders)
- `src/tools/calendar-tools.ts` — 3 calendar tools (list-events, get-event, create-event)
- `src/types.ts` — Shared interfaces: `OutlookMessage`, `OutlookEvent`, `OutlookFolder`

**AppleScript patterns:**
- All Outlook interaction happens through embedded AppleScript strings
- `asString()` must be used to safely embed user-provided strings into AppleScript (escapes quotes, newlines, tabs)
- `flattenForTSV()` strips special characters from values used as TSV fields
- TSV is the serialization format between AppleScript output and TypeScript parsing
- Each tool call spawns a new `osascript` process (stateless, no persistent connection)

**Error handling:**
- AppleScript errors surface as strings prefixed with `"ERROR: "`
- Each tool wraps its logic in try/catch and returns `{ isError: true }` on failure
- Temp files are cleaned up in `finally` blocks
