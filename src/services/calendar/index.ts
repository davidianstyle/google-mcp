import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { google } from "googleapis";
import { z } from "zod";
import { ServiceContext } from "../../types.js";
import { textResult } from "../../utils/formatting.js";
import { mapGoogleError } from "../../utils/errors.js";

import { calendar_v3 } from "googleapis";

function formatEvent(e: calendar_v3.Schema$Event): Record<string, unknown> {
  return {
    id: e.id,
    summary: e.summary,
    description: e.description,
    start: e.start,
    end: e.end,
    location: e.location,
    status: e.status,
    htmlLink: e.htmlLink,
    attendees: e.attendees,
    conferenceData: e.conferenceData,
    recurrence: e.recurrence,
    creator: e.creator,
    organizer: e.organizer,
  };
}

function summarizeAttendee(a: calendar_v3.Schema$EventAttendee): { email: string | undefined; responseStatus: string | undefined } {
  return { email: a.email ?? undefined, responseStatus: a.responseStatus ?? undefined };
}

/**
 * List-context event formatter: always trims attendees down to
 * email+responseStatus (full attendee objects carry displayName, organizer/
 * optional/resource flags, comment, etc. that a list view rarely needs), and
 * omits conferenceData unless includeDetails is true (it's a sizeable nested
 * object — entry points, phone numbers, notes — that's mostly only useful
 * when looking at one specific event).
 */
export function formatEventForList(e: calendar_v3.Schema$Event, includeDetails: boolean): Record<string, unknown> {
  return {
    ...formatEvent(e),
    attendees: e.attendees?.map(summarizeAttendee),
    conferenceData: includeDetails ? e.conferenceData : undefined,
  };
}

// Matches a trailing "Z" or "+HH:MM"/"-HH:MM" UTC offset on an ISO 8601
// dateTime string, meaning the timestamp is already unambiguous and doesn't
// need a timeZone field to be interpreted correctly.
const HAS_OFFSET_PATTERN = /(?:[Zz]|[+-]\d{2}:\d{2})$/;
function hasExplicitOffset(dateTime: string): boolean {
  return HAS_OFFSET_PATTERN.test(dateTime);
}

// Caches each calendar's default timeZone for the lifetime of the process,
// so defaulting a timed event to "the calendar's own timezone" costs at
// most one calendars.get per calendarId rather than one per create call.
const calendarTimeZoneCache = new Map<string, string>();
async function getCalendarTimeZone(cal: calendar_v3.Calendar, calendarId: string): Promise<string | undefined> {
  const cached = calendarTimeZoneCache.get(calendarId);
  if (cached) return cached;
  const res = await cal.calendars.get({ calendarId, fields: "timeZone" });
  const timeZone = res.data.timeZone || undefined;
  if (timeZone) calendarTimeZoneCache.set(calendarId, timeZone);
  return timeZone;
}

// Caches the authenticated user's own primary-calendar email for the
// lifetime of the process. calendar_respond_to_event only needs this as a
// fallback (an event where the user isn't flagged `self` on any attendee),
// and the answer can't change mid-process, so one calendarList.get call
// covers every subsequent respond call. Caches a failed/ungranted lookup too
// (as undefined) rather than re-attempting a doomed call every time.
let cachedPrimaryEmail: string | undefined;
let primaryEmailFetched = false;
async function getPrimaryEmail(cal: calendar_v3.Calendar): Promise<string | undefined> {
  if (primaryEmailFetched) return cachedPrimaryEmail;
  try {
    const primary = await cal.calendarList.get({ calendarId: "primary" });
    cachedPrimaryEmail = primary.data.id?.toLowerCase() || undefined;
  } catch {
    // calendarList scope not granted; cache the miss so we don't retry every call.
    cachedPrimaryEmail = undefined;
  }
  primaryEmailFetched = true;
  return cachedPrimaryEmail;
}

const ALL_DAY_END_NOTE = "For all-day events, use a 'YYYY-MM-DD' date (no time). The end date is exclusive — a single-day event on 2026-06-22 has start '2026-06-22' and end '2026-06-23'.";

export function registerCalendarTools(server: McpServer, ctx: ServiceContext): void {
  const api = google.calendar({ version: "v3", auth: ctx.auth });

  server.tool("calendar_list_events", "List events from a calendar within a time range", {
    calendarId: z.string().describe("Calendar ID (use 'primary' for main calendar)"),
    timeMin: z.string().optional().describe("Start of time range (ISO 8601)"),
    timeMax: z.string().optional().describe("End of time range (ISO 8601)"),
    timeZone: z.string().optional().describe("IANA timezone (e.g., 'America/New_York')"),
    maxResults: z.number().optional().describe("Maximum events to return"),
    pageToken: z.string().optional().describe("Token from a previous call's nextPageToken to fetch the next page"),
    includeDetails: z.boolean().optional().default(false).describe("Include full conferenceData per event. Default false always trims attendees to email+responseStatus and omits conferenceData to keep list responses small."),
  }, async ({ calendarId, timeMin, timeMax, timeZone, maxResults, pageToken, includeDetails }) => {
    const res = await api.events.list({
      calendarId,
      timeMin, timeMax, timeZone,
      maxResults: maxResults || 50,
      pageToken,
      singleEvents: true,
      orderBy: "startTime",
    });
    return textResult({
      events: res.data.items?.map((e) => formatEventForList(e, includeDetails)) || [],
      nextPageToken: res.data.nextPageToken,
    });
  });

  server.tool("calendar_create_event", "Create a new calendar event", {
    calendarId: z.string().describe("Calendar ID (use 'primary' for main calendar)"),
    summary: z.string().describe("Event title"),
    start: z.string().describe(`Start time (ISO 8601 datetime or date for all-day). ${ALL_DAY_END_NOTE}`),
    end: z.string().describe(`End time (ISO 8601 datetime or date for all-day). ${ALL_DAY_END_NOTE}`),
    description: z.string().optional(),
    location: z.string().optional(),
    attendees: z.array(z.object({ email: z.string(), displayName: z.string().optional(), optional: z.boolean().optional() })).optional(),
    timeZone: z.string().optional().describe("IANA timezone for a timed event. If omitted and start/end don't carry a UTC offset, defaults to the calendar's own timezone."),
    recurrence: z.array(z.string()).optional().describe("RFC5545 recurrence rules"),
    conferenceData: z.object({
      createRequest: z.object({
        requestId: z.string(),
        conferenceSolutionKey: z.object({ type: z.enum(["hangoutsMeet", "eventHangout", "eventNamedHangout", "addOn"]) }),
      }),
    }).optional(),
    reminders: z.object({
      useDefault: z.boolean(),
      overrides: z.array(z.object({ method: z.enum(["email", "popup"]).default("popup"), minutes: z.number() })).optional(),
    }).optional(),
    sendUpdates: z.enum(["all", "externalOnly", "none"]).optional(),
    visibility: z.enum(["default", "public", "private", "confidential"]).optional(),
    transparency: z.enum(["opaque", "transparent"]).optional(),
    colorId: z.string().optional(),
  }, async (opts) => {
    const cal = api;
    const isAllDay = !opts.start.includes("T");
    let timeZone = opts.timeZone;
    if (!isAllDay && !timeZone && !hasExplicitOffset(opts.start) && !hasExplicitOffset(opts.end)) {
      timeZone = await getCalendarTimeZone(cal, opts.calendarId);
    }
    const startField = isAllDay ? { date: opts.start } : { dateTime: opts.start, timeZone };
    const endField = isAllDay ? { date: opts.end } : { dateTime: opts.end, timeZone };

    const res = await cal.events.insert({
      calendarId: opts.calendarId,
      conferenceDataVersion: opts.conferenceData ? 1 : undefined,
      sendUpdates: opts.sendUpdates,
      requestBody: {
        summary: opts.summary,
        description: opts.description,
        location: opts.location,
        start: startField,
        end: endField,
        attendees: opts.attendees,
        recurrence: opts.recurrence,
        conferenceData: opts.conferenceData as unknown as undefined,
        reminders: opts.reminders,
        visibility: opts.visibility,
        transparency: opts.transparency,
        colorId: opts.colorId,
      },
    });
    return textResult(formatEvent(res.data ));
  });

  server.tool("calendar_create_events", "Create multiple calendar events at once", {
    calendarId: z.string(),
    events: z.array(z.object({
      summary: z.string(),
      start: z.string(),
      end: z.string(),
      description: z.string().optional(),
      location: z.string().optional(),
      attendees: z.array(z.object({ email: z.string() })).optional(),
      timeZone: z.string().optional(),
    })),
  }, async ({ calendarId, events }) => {
    const cal = api;
    const settled = await Promise.allSettled(events.map(async (evt) => {
      const isAllDay = !evt.start.includes("T");
      const res = await cal.events.insert({
        calendarId,
        requestBody: {
          summary: evt.summary,
          description: evt.description,
          location: evt.location,
          start: isAllDay ? { date: evt.start } : { dateTime: evt.start, timeZone: evt.timeZone },
          end: isAllDay ? { date: evt.end } : { dateTime: evt.end, timeZone: evt.timeZone },
          attendees: evt.attendees,
        },
      });
      return { id: res.data.id, summary: res.data.summary, htmlLink: res.data.htmlLink };
    }));

    const created: Array<{ id?: string | null; summary?: string | null; htmlLink?: string | null }> = [];
    const failed: Array<{ summary: string; error: string }> = [];
    settled.forEach((result, i) => {
      if (result.status === "fulfilled") {
        created.push(result.value);
      } else {
        failed.push({ summary: events[i].summary, error: mapGoogleError(result.reason) });
      }
    });

    return textResult(failed.length ? { created, failed } : created);
  });

  server.tool("calendar_get_event", "Get details of a specific event", {
    calendarId: z.string(),
    eventId: z.string(),
  }, async ({ calendarId, eventId }) => {
    const res = await api.events.get({ calendarId, eventId });
    return textResult(formatEvent(res.data ));
  });

  server.tool("calendar_update_event", "Update an existing calendar event", {
    calendarId: z.string(),
    eventId: z.string(),
    summary: z.string().optional(),
    start: z.string().optional().describe(`New start time (ISO 8601 datetime or date for all-day). ${ALL_DAY_END_NOTE}`),
    end: z.string().optional().describe(`New end time (ISO 8601 datetime or date for all-day). ${ALL_DAY_END_NOTE}`),
    description: z.string().optional(),
    location: z.string().optional(),
    attendees: z.array(z.object({ email: z.string(), displayName: z.string().optional() })).optional(),
    timeZone: z.string().optional().describe("IANA timezone to apply to a new start/end. If omitted, the event's existing timezone is preserved."),
    sendUpdates: z.enum(["all", "externalOnly", "none"]).optional(),
    colorId: z.string().optional(),
  }, async (opts) => {
    const cal = api;
    const requestBody: calendar_v3.Schema$Event = {};

    if (opts.summary !== undefined) requestBody.summary = opts.summary;
    if (opts.description !== undefined) requestBody.description = opts.description;
    if (opts.location !== undefined) requestBody.location = opts.location;
    if (opts.attendees !== undefined) requestBody.attendees = opts.attendees;
    if (opts.colorId !== undefined) requestBody.colorId = opts.colorId;

    if (opts.start !== undefined || opts.end !== undefined) {
      // Only fetch the existing event's timezone if we actually need it: a
      // timed start/end is being changed, no explicit timeZone was given,
      // and the datetime string itself doesn't already carry a UTC offset.
      const needsExistingTimeZone = !opts.timeZone && (
        (opts.start !== undefined && opts.start.includes("T") && !hasExplicitOffset(opts.start)) ||
        (opts.end !== undefined && opts.end.includes("T") && !hasExplicitOffset(opts.end))
      );
      let existingTimeZone: string | undefined;
      if (needsExistingTimeZone) {
        const existing = await cal.events.get({
          calendarId: opts.calendarId, eventId: opts.eventId,
          fields: "start(timeZone),end(timeZone)",
        });
        existingTimeZone = existing.data.start?.timeZone || existing.data.end?.timeZone || undefined;
      }

      if (opts.start !== undefined) {
        const isAllDay = !opts.start.includes("T");
        requestBody.start = isAllDay ? { date: opts.start } : { dateTime: opts.start, timeZone: opts.timeZone || existingTimeZone };
      }
      if (opts.end !== undefined) {
        const isAllDay = !opts.end.includes("T");
        requestBody.end = isAllDay ? { date: opts.end } : { dateTime: opts.end, timeZone: opts.timeZone || existingTimeZone };
      }
    }

    const res = await cal.events.patch({
      calendarId: opts.calendarId, eventId: opts.eventId,
      sendUpdates: opts.sendUpdates,
      requestBody,
    });
    return textResult(formatEvent(res.data ));
  });

  server.tool("calendar_delete_event", "Delete a calendar event", {
    calendarId: z.string(),
    eventId: z.string(),
    sendUpdates: z.enum(["all", "externalOnly", "none"]).optional(),
  }, async ({ calendarId, eventId, sendUpdates }) => {
    await api.events.delete({ calendarId, eventId, sendUpdates });
    return textResult({ success: true, eventId });
  });

  server.tool("calendar_search_events", "Search for events by text query", {
    calendarId: z.string(),
    query: z.string().describe("Free text search terms"),
    timeMin: z.string().optional(),
    timeMax: z.string().optional(),
    pageToken: z.string().optional().describe("Token from a previous call's nextPageToken to fetch the next page"),
  }, async ({ calendarId, query, timeMin, timeMax, pageToken }) => {
    const res = await api.events.list({
      calendarId, q: query, timeMin, timeMax,
      singleEvents: true, orderBy: "startTime", maxResults: 25,
      pageToken,
    });
    return textResult({
      events: res.data.items?.map(formatEvent) || [],
      nextPageToken: res.data.nextPageToken,
    });
  });

  server.tool("calendar_respond_to_event", "Respond to a calendar event invitation", {
    calendarId: z.string(),
    eventId: z.string(),
    responseStatus: z.enum(["accepted", "declined", "tentative"]),
    sendUpdates: z.enum(["all", "externalOnly", "none"]).optional(),
  }, async ({ calendarId, eventId, responseStatus, sendUpdates }) => {
    const cal = api;
    const existing = await cal.events.get({ calendarId, eventId });
    const attendees = existing.data.attendees || [];
    let me = attendees.find((a) => a.self);
    if (!me) {
      const myEmail = await getPrimaryEmail(cal);
      if (myEmail) {
        me = attendees.find((a) => a.email?.toLowerCase() === myEmail);
      }
    }
    if (!me) {
      throw new Error(`Cannot respond: authenticated user is not an attendee of event ${eventId}`);
    }
    me.responseStatus = responseStatus;

    const res = await cal.events.patch({
      calendarId, eventId, sendUpdates,
      requestBody: { attendees },
    });
    return textResult({ eventId: res.data.id, responseStatus });
  });

  server.tool("calendar_get_freebusy", "Check free/busy status for calendars", {
    timeMin: z.string().describe("Start of range (ISO 8601)"),
    timeMax: z.string().describe("End of range (ISO 8601)"),
    calendarIds: z.array(z.string()).describe("Calendar IDs to check"),
    timeZone: z.string().optional(),
  }, async ({ timeMin, timeMax, calendarIds, timeZone }) => {
    const res = await api.freebusy.query({
      requestBody: {
        timeMin, timeMax, timeZone,
        items: calendarIds.map((id) => ({ id })),
      },
    });
    return textResult(res.data.calendars);
  });

  server.tool("calendar_list_calendars", "List all calendars", {}, async () => {
    const res = await api.calendarList.list();
    return textResult(res.data.items?.map((c) => ({
      id: c.id, summary: c.summary, primary: c.primary, accessRole: c.accessRole, timeZone: c.timeZone,
    })) || []);
  });

  server.tool("calendar_list_colors", "List available event and calendar colors", {}, async () => {
    const res = await api.colors.get();
    return textResult({ event: res.data.event, calendar: res.data.calendar });
  });

  server.tool("calendar_get_current_time", "Get the current time in a specified timezone", {
    timeZone: z.string().optional().describe("IANA timezone (defaults to UTC)"),
  }, async ({ timeZone }) => {
    const now = new Date();
    const formatted = now.toLocaleString("en-US", { timeZone: timeZone || "UTC", dateStyle: "full", timeStyle: "long" });
    return textResult({ iso: now.toISOString(), formatted, timeZone: timeZone || "UTC" });
  });
}
