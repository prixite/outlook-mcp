import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { runAppleScript, asString, parseTSV } from '../applescript.js';
import type { OutlookMessage, OutlookFolder } from '../types.js';

function formatError(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

// ---------------------------------------------------------------------------
// AppleScript helpers
// ---------------------------------------------------------------------------

/**
 * AppleScript snippet that sets `targetFolder` based on a folder name string.
 * The variables `folderParam` and `accountFilter` must be set before this runs.
 * If `accountFilter` is non-empty, only folders belonging to that account
 * (matched by email address or account name) are considered.
 * For Inbox with no account filter, prefers the account with the most unread.
 */
const FOLDER_SELECTOR = `
  set targetFolder to missing value
  set searchName to ""
  if folderParam is "sent" or folderParam is "Sent Items" then
    set searchName to "Sent Items"
  else if folderParam is "drafts" or folderParam is "Drafts" then
    set searchName to "Drafts"
  else if folderParam is "deleted" or folderParam is "Deleted Items" or folderParam is "Trash" then
    set searchName to "Deleted Items"
  else if folderParam is not "" and folderParam is not "inbox" and folderParam is not "Inbox" then
    set searchName to folderParam
  else
    set searchName to "Inbox"
  end if

  set bestUnread to -1
  repeat with f in mail folders
    if (name of f) is searchName then
      set accountOk to true
      if accountFilter is not "" then
        set accountOk to false
        try
          set folderAcct to account of f
          try
            if (email address of folderAcct) is accountFilter then set accountOk to true
          end try
          if not accountOk then
            try
              if (name of folderAcct) is accountFilter then set accountOk to true
            end try
          end if
        end try
      end if
      if accountOk then
        if searchName is "Inbox" and accountFilter is "" then
          set fUnread to 0
          try
            set fUnread to unread count of f
          end try
          if fUnread > bestUnread then
            set bestUnread to fUnread
            set targetFolder to f
          end if
        else if accountOk then
          if targetFolder is missing value then
            set targetFolder to f
          end if
        end if
      end if
    end if
  end repeat

  if targetFolder is missing value then
    if accountFilter is "" then
      try
        set targetFolder to inbox
      end try
    else
      return "ERROR: No folder found for account " & accountFilter
    end if
  end if
`;

/** AppleScript handler for replacing text (defined at script level). */
const REPLACE_HANDLER = `
on replaceChars(theText, badChars)
  set AppleScript's text item delimiters to badChars
  set theItems to every text item of theText
  set AppleScript's text item delimiters to " "
  set cleaned to theItems as string
  set AppleScript's text item delimiters to ""
  return cleaned
end replaceChars
`;

/**
 * AppleScript snippet that formats a message reference `msg` into a TSV line
 * and appends it to `output`.
 * Fields: id, subject, senderName, senderEmail, dateReceived, isRead, hasAttachments, preview(150)
 */
const FORMAT_MESSAGE = `
  set msgId to id of msg as string

  set msgSubject to ""
  try
    set msgSubject to subject of msg
    if msgSubject is missing value then set msgSubject to ""
    set msgSubject to my replaceChars(msgSubject, tab)
  end try

  set senderName to ""
  set senderEmail to ""
  try
    set senderObj to sender of msg
    set senderName to name of senderObj
    if senderName is missing value then set senderName to ""
    set senderEmail to address of senderObj
    if senderEmail is missing value then set senderEmail to ""
  end try

  set msgDate to ""
  try
    set msgDate to (time received of msg) as string
  end try

  set msgIsRead to "false"
  try
    if is read of msg then set msgIsRead to "true"
  end try

  set hasAttach to "false"
  try
    if (count of attachments of msg) > 0 then set hasAttach to "true"
  end try

  set preview to ""
  try
    set fullContent to plain text content of msg
    if fullContent is missing value then set fullContent to ""
    if length of fullContent > 150 then
      set preview to (text 1 thru 150 of fullContent)
    else
      set preview to fullContent
    end if
    set preview to my replaceChars(preview, tab)
    set preview to my replaceChars(preview, (character id 10))
    set preview to my replaceChars(preview, (character id 13))
  end try

  set output to output & msgId & tab & msgSubject & tab & senderName & tab & senderEmail & tab & msgDate & tab & msgIsRead & tab & hasAttach & tab & preview & linefeed
`;

/** Parse TSV rows into OutlookMessage objects (no body). */
function parseMessageRows(rows: string[][], folder?: string): OutlookMessage[] {
  return rows
    .filter((r) => r.length >= 7)
    .map((r) => ({
      id: parseInt(r[0], 10),
      subject: r[1] ?? '',
      senderName: r[2] ?? '',
      senderEmail: r[3] ?? '',
      dateReceived: r[4] ?? '',
      isRead: r[5] === 'true',
      hasAttachments: r[6] === 'true',
      folder: folder ?? 'Inbox',
      preview: r[7] ?? '',
    }));
}

// ---------------------------------------------------------------------------
// Tool registration
// ---------------------------------------------------------------------------

export function registerEmailTools(server: McpServer): void {
  // ─── List Messages ────────────────────────────────────────────────────────
  server.registerTool(
    'outlook_list_messages',
    {
      title: 'List Outlook Messages',
      description:
        'List recent email messages from an Outlook folder (default: inbox). Returns subject, sender, date, read status, and a short preview.',
      inputSchema: {
        folder: z
          .string()
          .optional()
          .default('inbox')
          .describe('Folder name: "inbox", "sent", "drafts", "deleted", or a custom folder name'),
        account: z
          .string()
          .optional()
          .default('')
          .describe('Account email to filter by (e.g. "you@toptal.com"). Use outlook_list_accounts to see available accounts.'),
        limit: z
          .number()
          .int()
          .min(1)
          .max(100)
          .optional()
          .default(20)
          .describe('Max number of messages to return (default 20, max 100)'),
        unread_only: z
          .boolean()
          .optional()
          .default(false)
          .describe('If true, return only unread messages'),
      },
    },
    async ({ folder, account, limit, unread_only }) => {
      const folderParam = folder ?? 'inbox';
      const accountFilter = account ?? '';
      const limitNum = limit ?? 20;
      const unreadOnly = unread_only ?? false;

      const script = `
${REPLACE_HANDLER}

tell application "Microsoft Outlook"
  set folderParam to ${asString(folderParam)}
  set accountFilter to ${asString(accountFilter)}
  ${FOLDER_SELECTOR}

  set output to ""
  set msgCount to 0
  set allMsgs to messages of targetFolder

  repeat with msg in allMsgs
    if msgCount >= ${limitNum} then exit repeat

    -- unread filter
    if ${unreadOnly} then
      try
        if is read of msg then
          -- skip read messages
        else
          ${FORMAT_MESSAGE}
          set msgCount to msgCount + 1
        end if
      end try
    else
      try
        ${FORMAT_MESSAGE}
        set msgCount to msgCount + 1
      end try
    end if
  end repeat

  return output
end tell
`;
      try {
        const raw = await runAppleScript(script);
        if (raw.startsWith('ERROR:')) {
          return { content: [{ type: 'text' as const, text: raw }], isError: true };
        }
        const messages = parseMessageRows(parseTSV(raw), folderParam);
        if (messages.length === 0) {
          return {
            content: [{ type: 'text' as const, text: `No messages found in "${folderParam}".` }],
          };
        }
        return {
          content: [{ type: 'text' as const, text: JSON.stringify(messages, null, 2) }],
        };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── Read Message ─────────────────────────────────────────────────────────
  server.registerTool(
    'outlook_read_message',
    {
      title: 'Read Outlook Message',
      description:
        'Read the full content of an Outlook email message by its ID. Searches inbox, sent, and drafts.',
      inputSchema: {
        message_id: z
          .number()
          .int()
          .positive()
          .describe(
            'The integer message ID returned by outlook_list_messages or outlook_search_messages'
          ),
      },
    },
    async ({ message_id }) => {
      const script = `
${REPLACE_HANDLER}

tell application "Microsoft Outlook"
  set targetMsg to missing value
  set targetFolderName to ""

  -- Search all mail folders
  repeat with f in mail folders
    try
      set targetMsg to (first message of f whose id = ${message_id})
      set targetFolderName to name of f
      exit repeat
    end try
  end repeat

  if targetMsg is missing value then
    return "ERROR: Message not found"
  end if

  set msg to targetMsg
  set msgId to id of msg as string

  set msgSubject to ""
  try
    set msgSubject to subject of msg
    if msgSubject is missing value then set msgSubject to ""
  end try

  set senderName to ""
  set senderEmail to ""
  try
    set senderObj to sender of msg
    set senderName to name of senderObj
    if senderName is missing value then set senderName to ""
    set senderEmail to address of senderObj
    if senderEmail is missing value then set senderEmail to ""
  end try

  set toList to ""
  try
    set toRecips to to recipients of msg
    repeat with r in toRecips
      set rName to ""
      set rEmail to ""
      try
        set rName to name of r
        if rName is missing value then set rName to ""
      end try
      try
        set rAddr to email address of r
        set rEmail to address of rAddr
        if rEmail is missing value then set rEmail to ""
      end try
      if toList is "" then
        set toList to rName & " <" & rEmail & ">"
      else
        set toList to toList & ", " & rName & " <" & rEmail & ">"
      end if
    end repeat
  end try

  set msgDate to ""
  try
    set msgDate to (time received of msg) as string
  end try

  set msgIsRead to "false"
  try
    if is read of msg then set msgIsRead to "true"
  end try

  set hasAttach to "false"
  try
    if (count of attachments of msg) > 0 then set hasAttach to "true"
  end try

  set msgBody to ""
  try
    set msgBody to plain text content of msg
    if msgBody is missing value then set msgBody to ""
  end try

  -- Output: header fields tab-separated on first line, then BODY: marker, then body
  set header to msgId & tab & msgSubject & tab & senderName & tab & senderEmail & tab & toList & tab & msgDate & tab & msgIsRead & tab & hasAttach & tab & targetFolderName
  return header & linefeed & "BODY:" & linefeed & msgBody
end tell
`;
      try {
        const raw = await runAppleScript(script);
        if (raw.startsWith('ERROR:')) {
          return { content: [{ type: 'text' as const, text: raw }], isError: true };
        }

        const bodyMarkerIndex = raw.indexOf('\nBODY:\n');
        if (bodyMarkerIndex === -1) {
          return { content: [{ type: 'text' as const, text: raw }] };
        }

        const headerLine = raw.slice(0, bodyMarkerIndex);
        const body = raw.slice(bodyMarkerIndex + '\nBODY:\n'.length);
        const fields = headerLine.split('\t');

        const message: OutlookMessage = {
          id: parseInt(fields[0] ?? '0', 10),
          subject: fields[1] ?? '',
          senderName: fields[2] ?? '',
          senderEmail: fields[3] ?? '',
          dateReceived: fields[5] ?? '',
          isRead: fields[6] === 'true',
          hasAttachments: fields[7] === 'true',
          folder: fields[8] ?? '',
          body,
        };

        const toField = fields[4] ?? '';
        const output = {
          ...message,
          to: toField,
        };

        return { content: [{ type: 'text' as const, text: JSON.stringify(output, null, 2) }] };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── Search Messages ──────────────────────────────────────────────────────
  server.registerTool(
    'outlook_search_messages',
    {
      title: 'Search Outlook Messages',
      description:
        'Search recent inbox messages by keyword (matches subject, sender name, or sender email).',
      inputSchema: {
        query: z
          .string()
          .min(1)
          .describe('Search keyword to match against subject, sender name, or sender email'),
        account: z
          .string()
          .optional()
          .default('')
          .describe('Account email to filter by (e.g. "you@toptal.com"). Use outlook_list_accounts to see available accounts.'),
        limit: z
          .number()
          .int()
          .min(1)
          .max(50)
          .optional()
          .default(10)
          .describe('Max results to return (default 10)'),
        scan_count: z
          .number()
          .int()
          .min(10)
          .max(500)
          .optional()
          .default(100)
          .describe('How many recent inbox messages to scan (default 100)'),
      },
    },
    async ({ query, account, limit, scan_count }) => {
      const scanCount = scan_count ?? 100;
      const maxResults = limit ?? 10;
      const accountFilter = account ?? '';
      const lowerQuery = query.toLowerCase();

      // Fetch recent messages from inbox, filter client-side
      const script = `
${REPLACE_HANDLER}

tell application "Microsoft Outlook"
  set output to ""
  set msgCount to 0
  set accountFilter to ${asString(accountFilter)}

  -- Find inbox; if accountFilter set, match that account; otherwise pick most unread
  set activeInbox to missing value
  set bestUnread to -1
  repeat with f in mail folders
    if (name of f) is "Inbox" then
      set accountOk to true
      if accountFilter is not "" then
        set accountOk to false
        try
          set folderAcct to account of f
          try
            if (email address of folderAcct) is accountFilter then set accountOk to true
          end try
          if not accountOk then
            try
              if (name of folderAcct) is accountFilter then set accountOk to true
            end try
          end if
        end try
      end if
      if accountOk then
        if accountFilter is not "" then
          set activeInbox to f
          exit repeat
        else
          set fUnread to 0
          try
            set fUnread to unread count of f
          end try
          if fUnread > bestUnread then
            set bestUnread to fUnread
            set activeInbox to f
          end if
        end if
      end if
    end if
  end repeat
  if activeInbox is missing value then
    if accountFilter is not "" then
      return "ERROR: No inbox found for account " & accountFilter
    end if
    try
      set activeInbox to inbox
    end try
  end if

  set allMsgs to messages of activeInbox

  repeat with msg in allMsgs
    if msgCount >= ${scanCount} then exit repeat
    try
      ${FORMAT_MESSAGE}
      set msgCount to msgCount + 1
    end try
  end repeat

  return output
end tell
`;
      try {
        const raw = await runAppleScript(script);
        if (raw.startsWith('ERROR:')) {
          return { content: [{ type: 'text' as const, text: raw }], isError: true };
        }
        const allMessages = parseMessageRows(parseTSV(raw), 'Inbox');

        const matches = allMessages
          .filter(
            (m) =>
              m.subject.toLowerCase().includes(lowerQuery) ||
              m.senderName.toLowerCase().includes(lowerQuery) ||
              m.senderEmail.toLowerCase().includes(lowerQuery) ||
              (m.preview ?? '').toLowerCase().includes(lowerQuery)
          )
          .slice(0, maxResults);

        if (matches.length === 0) {
          return {
            content: [{ type: 'text' as const, text: `No messages found matching "${query}".` }],
          };
        }
        return { content: [{ type: 'text' as const, text: JSON.stringify(matches, null, 2) }] };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── Send Email ───────────────────────────────────────────────────────────
  server.registerTool(
    'outlook_send_email',
    {
      title: 'Send Email via Outlook',
      description: 'Compose and send an email from Outlook.',
      inputSchema: {
        to: z.union([z.string(), z.array(z.string())]).describe('Recipient email address(es)'),
        subject: z.string().describe('Email subject'),
        body: z.string().describe('Email body text'),
        cc: z
          .union([z.string(), z.array(z.string())])
          .optional()
          .describe('CC recipient(s)'),
      },
    },
    async ({ to, subject, body, cc }) => {
      const toList = Array.isArray(to) ? to : [to];
      const ccList = cc ? (Array.isArray(cc) ? cc : [cc]) : [];

      // Build recipient AppleScript lines
      const toLines = toList
        .map(
          (addr) =>
            `make new to recipient at newMsg with properties {email address: {address: ${asString(addr)}, name: ""}}`
        )
        .join('\n  ');
      const ccLines = ccList
        .map(
          (addr) =>
            `make new cc recipient at newMsg with properties {email address: {address: ${asString(addr)}, name: ""}}`
        )
        .join('\n  ');

      const script = `
tell application "Microsoft Outlook"
  set newMsg to make new outgoing message with properties {subject: ${asString(subject)}, content: ${asString(body)}}
  ${toLines}
  ${ccLines}
  send newMsg
end tell
return "sent"
`;
      try {
        await runAppleScript(script);
        return {
          content: [
            { type: 'text' as const, text: `Email sent successfully to: ${toList.join(', ')}` },
          ],
        };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── Reply to Message ─────────────────────────────────────────────────────
  server.registerTool(
    'outlook_reply',
    {
      title: 'Reply to Outlook Message',
      description: 'Reply to an email message by its ID.',
      inputSchema: {
        message_id: z.number().int().positive().describe('ID of the message to reply to'),
        body: z.string().describe('Reply body text (prepended before the quoted original)'),
        reply_all: z
          .boolean()
          .optional()
          .default(false)
          .describe('If true, reply to all recipients'),
      },
    },
    async ({ message_id, body, reply_all }) => {
      const replyAllFlag = reply_all ? 'true' : 'false';
      const script = `
tell application "Microsoft Outlook"
  set targetMsg to missing value

  repeat with f in mail folders
    try
      set targetMsg to (first message of f whose id = ${message_id})
      exit repeat
    end try
  end repeat

  if targetMsg is missing value then
    return "ERROR: Message not found"
  end if

  set replyMsg to reply targetMsg reply all ${replyAllFlag}
  set content of replyMsg to ${asString(body)} & return & return & content of replyMsg
  send replyMsg
  return "replied"
end tell
`;
      try {
        const result = await runAppleScript(script);
        if (result.startsWith('ERROR:')) {
          return { content: [{ type: 'text' as const, text: result }], isError: true };
        }
        return { content: [{ type: 'text' as const, text: `Reply sent successfully.` }] };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── Forward Message ──────────────────────────────────────────────────────
  server.registerTool(
    'outlook_forward',
    {
      title: 'Forward Outlook Message',
      description: 'Forward an email message to one or more recipients.',
      inputSchema: {
        message_id: z.number().int().positive().describe('ID of the message to forward'),
        to: z.union([z.string(), z.array(z.string())]).describe('Recipient email address(es)'),
        body: z
          .string()
          .optional()
          .describe('Optional text to prepend before the forwarded message'),
      },
    },
    async ({ message_id, to, body }) => {
      const toList = Array.isArray(to) ? to : [to];
      const toLines = toList
        .map(
          (addr) =>
            `make new to recipient at fwdMsg with properties {email address: {address: ${asString(addr)}, name: ""}}`
        )
        .join('\n  ');
      const bodyPrefix = body
        ? `set content of fwdMsg to ${asString(body)} & return & return & content of fwdMsg`
        : '';

      const script = `
tell application "Microsoft Outlook"
  set targetMsg to missing value

  repeat with f in mail folders
    try
      set targetMsg to (first message of f whose id = ${message_id})
      exit repeat
    end try
  end repeat

  if targetMsg is missing value then
    return "ERROR: Message not found"
  end if

  set fwdMsg to forward targetMsg
  ${toLines}
  ${bodyPrefix}
  send fwdMsg
  return "forwarded"
end tell
`;
      try {
        const result = await runAppleScript(script);
        if (result.startsWith('ERROR:')) {
          return { content: [{ type: 'text' as const, text: result }], isError: true };
        }
        return {
          content: [{ type: 'text' as const, text: `Message forwarded to: ${toList.join(', ')}` }],
        };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── Delete Message ───────────────────────────────────────────────────────
  server.registerTool(
    'outlook_delete_message',
    {
      title: 'Delete Outlook Message',
      description: 'Move an email message to the Deleted Items folder.',
      inputSchema: {
        message_id: z.number().int().positive().describe('ID of the message to delete'),
      },
    },
    async ({ message_id }) => {
      const script = `
tell application "Microsoft Outlook"
  set targetMsg to missing value

  repeat with f in mail folders
    try
      set targetMsg to (first message of f whose id = ${message_id})
      exit repeat
    end try
  end repeat

  if targetMsg is missing value then
    return "ERROR: Message not found"
  end if

  delete targetMsg
  return "deleted"
end tell
`;
      try {
        const result = await runAppleScript(script);
        if (result.startsWith('ERROR:')) {
          return { content: [{ type: 'text' as const, text: result }], isError: true };
        }
        return {
          content: [
            { type: 'text' as const, text: `Message ${message_id} moved to Deleted Items.` },
          ],
        };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── Mark as Read ─────────────────────────────────────────────────────────
  server.registerTool(
    'outlook_mark_read',
    {
      title: 'Mark Outlook Message as Read/Unread',
      description: 'Mark an email message as read or unread.',
      inputSchema: {
        message_id: z.number().int().positive().describe('ID of the message'),
        read: z
          .boolean()
          .optional()
          .default(true)
          .describe('true = mark as read, false = mark as unread'),
      },
    },
    async ({ message_id, read }) => {
      const readValue = (read ?? true) ? 'true' : 'false';
      const script = `
tell application "Microsoft Outlook"
  set targetMsg to missing value

  repeat with f in mail folders
    try
      set targetMsg to (first message of f whose id = ${message_id})
      exit repeat
    end try
  end repeat

  if targetMsg is missing value then
    return "ERROR: Message not found"
  end if

  set is read of targetMsg to ${readValue}
  return "done"
end tell
`;
      try {
        const result = await runAppleScript(script);
        if (result.startsWith('ERROR:')) {
          return { content: [{ type: 'text' as const, text: result }], isError: true };
        }
        const label = (read ?? true) ? 'read' : 'unread';
        return {
          content: [{ type: 'text' as const, text: `Message ${message_id} marked as ${label}.` }],
        };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── List Accounts ────────────────────────────────────────────────────────
  server.registerTool(
    'outlook_list_accounts',
    {
      title: 'List Outlook Accounts',
      description:
        'List all email accounts configured in Outlook, including Exchange, IMAP, and POP accounts.',
      inputSchema: {},
    },
    async () => {
      // Derive accounts by inspecting the first message in each Inbox folder.
      // exchange/imap/pop account collections are empty in modern Outlook for Mac.
      const script = `
tell application "Microsoft Outlook"
  set output to ""
  set seenNames to {}

  repeat with f in mail folders
    if (name of f) is "Inbox" then
      set displayName to ""

      -- Try to get account directly from the folder (works even for empty inboxes)
      try
        set acct to account of f
        try
          set acctEmail to email address of acct
          if acctEmail is not missing value and acctEmail is not "" then
            set displayName to acctEmail
          end if
        end try
        if displayName is "" then
          try
            set acctName to name of acct
            if acctName is not missing value then set displayName to acctName
          end try
        end if
      end try

      -- Fallback: derive from first message if folder approach failed
      if displayName is "" then
        set msgList to every message of f
        if (count of msgList) > 0 then
          try
            set acct to account of (item 1 of msgList)
            try
              set acctEmail to email address of acct
              if acctEmail is not missing value and acctEmail is not "" then
                set displayName to acctEmail
              end if
            end try
            if displayName is "" then
              try
                set acctName to name of acct
                if acctName is not missing value then set displayName to acctName
              end try
            end if
          end try
        end if
      end if

      if displayName is not "" then
        set alreadySeen to false
        repeat with seen in seenNames
          if seen is displayName then
            set alreadySeen to true
            exit repeat
          end if
        end repeat
        if not alreadySeen then
          set end of seenNames to displayName
          set unreadCnt to unread count of f
          set output to output & displayName & tab & (unreadCnt as string) & linefeed
        end if
      end if
    end if
  end repeat

  return output
end tell
`;
      try {
        const raw = await runAppleScript(script);
        const rows = parseTSV(raw);
        const accounts = rows
          .filter((r) => r.length >= 1)
          .map((r) => ({
            email: r[0] ?? '',
            inboxUnread: parseInt(r[1] ?? '0', 10),
          }));
        if (accounts.length === 0) {
          return { content: [{ type: 'text' as const, text: 'No accounts found.' }] };
        }
        return { content: [{ type: 'text' as const, text: JSON.stringify(accounts, null, 2) }] };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── List Folders ─────────────────────────────────────────────────────────
  server.registerTool(
    'outlook_list_folders',
    {
      title: 'List Outlook Mail Folders',
      description: 'List all mail folders in Outlook.',
      inputSchema: {},
    },
    async () => {
      const script = `
tell application "Microsoft Outlook"
  set output to ""
  repeat with f in mail folders
    set folderName to name of f
    set unreadCount to 0
    try
      set unreadCount to unread count of f
    end try
    set output to output & folderName & tab & (unreadCount as string) & linefeed
  end repeat
  return output
end tell
`;
      try {
        const raw = await runAppleScript(script);
        const rows = parseTSV(raw);
        const folders: OutlookFolder[] = rows.map((r) => ({
          name: r[0] ?? '',
          unreadCount: parseInt(r[1] ?? '0', 10),
        }));
        if (folders.length === 0) {
          return { content: [{ type: 'text' as const, text: 'No folders found.' }] };
        }
        return { content: [{ type: 'text' as const, text: JSON.stringify(folders, null, 2) }] };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );
}
