import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { runAppleScript, asString, parseTSV } from '../applescript.js';
import type { OutlookEvent } from '../types.js';

function formatError(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

/**
 * Build an AppleScript snippet that sets `theDate` to the given ISO datetime string.
 * Accepts "YYYY-MM-DDTHH:MM:SS" or "YYYY-MM-DD" (all-day).
 */
function buildDateScript(varName: string, iso: string): string {
  // Parse the ISO string in TypeScript so we don't have to do it in AppleScript
  const datePart = iso.slice(0, 10);
  const timePart = iso.length >= 19 ? iso.slice(11, 19) : '00:00:00';
  const [yr, mo, dy] = datePart.split('-').map(Number);
  const [hr, mn, se] = timePart.split(':').map(Number);

  return `
  set ${varName} to current date
  set year of ${varName} to ${yr}
  set month of ${varName} to ${mo}
  set day of ${varName} to ${dy}
  set time of ${varName} to ${hr * 3600 + mn * 60 + se}
`;
}

export function registerCalendarTools(server: McpServer): void {
  // ─── List Events ──────────────────────────────────────────────────────────
  server.registerTool(
    'outlook_list_events',
    {
      title: 'List Outlook Calendar Events',
      description:
        'List upcoming calendar events from Outlook. Returns events sorted by start time.',
      inputSchema: {
        days: z
          .number()
          .int()
          .min(1)
          .max(365)
          .optional()
          .default(7)
          .describe('Number of days ahead to look (default 7)'),
        limit: z
          .number()
          .int()
          .min(1)
          .max(100)
          .optional()
          .default(20)
          .describe('Max events to return (default 20)'),
      },
    },
    async ({ days, limit }) => {
      const daysAhead = days ?? 7;
      const limitNum = limit ?? 20;

      const script = `
tell application "Microsoft Outlook"
  set now to current date
  set cutoff to now + (${daysAhead} * 24 * 3600)
  set output to ""
  set evtCount to 0

  set allEvents to calendar events
  repeat with evt in allEvents
    if evtCount >= ${limitNum} then exit repeat
    try
      set evtStart to start time of evt
      if evtStart >= now and evtStart <= cutoff then
        set evtId to id of evt as string

        set evtSubject to ""
        try
          set evtSubject to subject of evt
          if evtSubject is missing value then set evtSubject to ""
        end try

        set evtEnd to ""
        try
          set evtEnd to end time of evt as string
        end try

        set evtLoc to ""
        try
          set evtLoc to location of evt
          if evtLoc is missing value then set evtLoc to ""
          -- strip tabs
          set AppleScript's text item delimiters to tab
          set parts to every text item of evtLoc
          set AppleScript's text item delimiters to " "
          set evtLoc to parts as string
          set AppleScript's text item delimiters to ""
        end try

        set evtAllDay to "false"
        try
          if all day event of evt then set evtAllDay to "true"
        end try

        set output to output & evtId & tab & evtSubject & tab & (evtStart as string) & tab & evtEnd & tab & evtLoc & tab & evtAllDay & linefeed
        set evtCount to evtCount + 1
      end if
    end try
  end repeat

  return output
end tell
`;
      try {
        const raw = await runAppleScript(script);
        const rows = parseTSV(raw);
        const events: OutlookEvent[] = rows
          .filter((r) => r.length >= 6)
          .map((r) => ({
            id: parseInt(r[0], 10),
            subject: r[1] ?? '',
            startTime: r[2] ?? '',
            endTime: r[3] ?? '',
            location: r[4] ?? '',
            isAllDay: r[5] === 'true',
          }));

        if (events.length === 0) {
          return {
            content: [
              { type: 'text' as const, text: `No events found in the next ${daysAhead} day(s).` },
            ],
          };
        }
        return { content: [{ type: 'text' as const, text: JSON.stringify(events, null, 2) }] };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── Get Event ────────────────────────────────────────────────────────────
  server.registerTool(
    'outlook_get_event',
    {
      title: 'Get Outlook Calendar Event',
      description:
        'Get full details of a calendar event by its ID, including the description/notes.',
      inputSchema: {
        event_id: z
          .number()
          .int()
          .positive()
          .describe('The integer event ID returned by outlook_list_events'),
      },
    },
    async ({ event_id }) => {
      const script = `
tell application "Microsoft Outlook"
  set targetEvt to missing value
  try
    set targetEvt to (first calendar event whose id = ${event_id})
  end try

  if targetEvt is missing value then
    return "ERROR: Event not found"
  end if

  set evt to targetEvt
  set evtId to id of evt as string

  set evtSubject to ""
  try
    set evtSubject to subject of evt
    if evtSubject is missing value then set evtSubject to ""
  end try

  set evtStart to ""
  try
    set evtStart to start time of evt as string
  end try

  set evtEnd to ""
  try
    set evtEnd to end time of evt as string
  end try

  set evtLoc to ""
  try
    set evtLoc to location of evt
    if evtLoc is missing value then set evtLoc to ""
  end try

  set evtAllDay to "false"
  try
    if all day event of evt then set evtAllDay to "true"
  end try

  set evtBody to ""
  try
    set evtBody to content of evt
    if evtBody is missing value then set evtBody to ""
  end try

  set header to evtId & tab & evtSubject & tab & evtStart & tab & evtEnd & tab & evtLoc & tab & evtAllDay
  return header & linefeed & "BODY:" & linefeed & evtBody
end tell
`;
      try {
        const raw = await runAppleScript(script);
        if (raw.startsWith('ERROR:')) {
          return { content: [{ type: 'text' as const, text: raw }], isError: true };
        }

        const bodyMarkerIndex = raw.indexOf('\nBODY:\n');
        const headerLine = bodyMarkerIndex >= 0 ? raw.slice(0, bodyMarkerIndex) : raw;
        const body = bodyMarkerIndex >= 0 ? raw.slice(bodyMarkerIndex + '\nBODY:\n'.length) : '';
        const fields = headerLine.split('\t');

        const event: OutlookEvent = {
          id: parseInt(fields[0] ?? '0', 10),
          subject: fields[1] ?? '',
          startTime: fields[2] ?? '',
          endTime: fields[3] ?? '',
          location: fields[4] ?? '',
          isAllDay: fields[5] === 'true',
          body,
        };

        return { content: [{ type: 'text' as const, text: JSON.stringify(event, null, 2) }] };
      } catch (err) {
        return {
          content: [{ type: 'text' as const, text: `Error: ${formatError(err)}` }],
          isError: true,
        };
      }
    }
  );

  // ─── Create Event ─────────────────────────────────────────────────────────
  server.registerTool(
    'outlook_create_event',
    {
      title: 'Create Outlook Calendar Event',
      description: 'Create a new calendar event in Outlook.',
      inputSchema: {
        subject: z.string().describe('Event title/subject'),
        start_datetime: z
          .string()
          .describe('Start date/time in ISO 8601 format: "YYYY-MM-DDTHH:MM:SS"'),
        end_datetime: z
          .string()
          .describe('End date/time in ISO 8601 format: "YYYY-MM-DDTHH:MM:SS"'),
        location: z.string().optional().describe('Event location'),
        body: z.string().optional().describe('Event description/notes'),
        all_day: z
          .boolean()
          .optional()
          .default(false)
          .describe('If true, create as an all-day event'),
      },
    },
    async ({ subject, start_datetime, end_datetime, location, body, all_day }) => {
      const startScript = buildDateScript('startDate', start_datetime);
      const endScript = buildDateScript('endDate', end_datetime);
      const locationLine = location ? `set location of newEvent to ${asString(location)}` : '';
      const bodyLine = body ? `set content of newEvent to ${asString(body)}` : '';
      const allDayLine = (all_day ?? false) ? `set all day event of newEvent to true` : '';

      const script = `
tell application "Microsoft Outlook"
  ${startScript}
  ${endScript}
  set newEvent to make new calendar event with properties {subject: ${asString(subject)}, start time: startDate, end time: endDate}
  ${locationLine}
  ${bodyLine}
  ${allDayLine}
  set evtId to id of newEvent as string
  return "created:" & evtId
end tell
`;
      try {
        const result = await runAppleScript(script);
        const idMatch = result.match(/^created:(\d+)$/);
        const idStr = idMatch ? ` (ID: ${idMatch[1]})` : '';
        return {
          content: [
            {
              type: 'text' as const,
              text: `Calendar event "${subject}" created successfully.${idStr}`,
            },
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
}
