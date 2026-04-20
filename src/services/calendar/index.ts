import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { google } from "googleapis";
import { z } from "zod";
import { ServiceContext } from "../../types.js";
import { textResult } from "../../utils/formatting.js";

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

export function registerCalendarTools(server: McpServer, ctx: ServiceContext): void {
  const api = () => google.calendar({ version: "v3", auth: ctx.auth });

  server.tool("calendar_list_events", "List events from a calendar within a time range", {
    calendarId: z.string().describe("Calendar ID (use 'primary' for main calendar)"),
    timeMin: z.string().optional().describe("Start of time range (ISO 8601)"),
    timeMax: z.string().optional().describe("End of time range (ISO 8601)"),
    timeZone: z.string().optional().describe("IANA timezone (e.g., 'America/New_York')"),
    maxResults: z.number().optional().describe("Maximum events to return"),
  }, async ({ calendarId, timeMin, timeMax, timeZone, maxResults }) => {
    const res = await api().events.list({
      calendarId,
      timeMin, timeMax, timeZone,
      maxResults: maxResults || 50,
      singleEvents: true,
      orderBy: "startTime",
    });
    return textResult(res.data.items?.map(formatEvent) || []);
  });

  server.tool("calendar_create_event", "Create a new calendar event", {
    calendarId: z.string().describe("Calendar ID (use 'primary' for main calendar)"),
    summary: z.string().describe("Event title"),
    start: z.string().describe("Start time (ISO 8601 datetime or date for all-day)"),
    end: z.string().describe("End time (ISO 8601 datetime or date for all-day)"),
    description: z.string().optional(),
    location: z.string().optional(),
    attendees: z.array(z.object({ email: z.string(), displayName: z.string().optional(), optional: z.boolean().optional() })).optional(),
    timeZone: z.string().optional(),
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
    const isAllDay = !opts.start.includes("T");
    const startField = isAllDay ? { date: opts.start } : { dateTime: opts.start, timeZone: opts.timeZone };
    const endField = isAllDay ? { date: opts.end } : { dateTime: opts.end, timeZone: opts.timeZone };

    const res = await api().events.insert({
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
    const cal = api();
    const results = await Promise.all(events.map(async (evt) => {
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
    return textResult(results);
  });

  server.tool("calendar_get_event", "Get details of a specific event", {
    calendarId: z.string(),
    eventId: z.string(),
  }, async ({ calendarId, eventId }) => {
    const res = await api().events.get({ calendarId, eventId });
    return textResult(formatEvent(res.data ));
  });

  server.tool("calendar_update_event", "Update an existing calendar event", {
    calendarId: z.string(),
    eventId: z.string(),
    summary: z.string().optional(),
    start: z.string().optional(),
    end: z.string().optional(),
    description: z.string().optional(),
    location: z.string().optional(),
    attendees: z.array(z.object({ email: z.string(), displayName: z.string().optional() })).optional(),
    timeZone: z.string().optional(),
    sendUpdates: z.enum(["all", "externalOnly", "none"]).optional(),
    colorId: z.string().optional(),
  }, async (opts) => {
    const cal = api();
    const existing = await cal.events.get({ calendarId: opts.calendarId, eventId: opts.eventId });
    const body = existing.data;

    if (opts.summary !== undefined) body.summary = opts.summary;
    if (opts.description !== undefined) body.description = opts.description;
    if (opts.location !== undefined) body.location = opts.location;
    if (opts.attendees !== undefined) body.attendees = opts.attendees;
    if (opts.colorId !== undefined) body.colorId = opts.colorId;
    if (opts.start) {
      const isAllDay = !opts.start.includes("T");
      body.start = isAllDay ? { date: opts.start } : { dateTime: opts.start, timeZone: opts.timeZone };
    }
    if (opts.end) {
      const isAllDay = !opts.end.includes("T");
      body.end = isAllDay ? { date: opts.end } : { dateTime: opts.end, timeZone: opts.timeZone };
    }

    const res = await cal.events.update({
      calendarId: opts.calendarId, eventId: opts.eventId,
      sendUpdates: opts.sendUpdates,
      requestBody: body,
    });
    return textResult(formatEvent(res.data ));
  });

  server.tool("calendar_delete_event", "Delete a calendar event", {
    calendarId: z.string(),
    eventId: z.string(),
    sendUpdates: z.enum(["all", "externalOnly", "none"]).optional(),
  }, async ({ calendarId, eventId, sendUpdates }) => {
    await api().events.delete({ calendarId, eventId, sendUpdates });
    return textResult({ success: true, eventId });
  });

  server.tool("calendar_search_events", "Search for events by text query", {
    calendarId: z.string(),
    query: z.string().describe("Free text search terms"),
    timeMin: z.string().optional(),
    timeMax: z.string().optional(),
  }, async ({ calendarId, query, timeMin, timeMax }) => {
    const res = await api().events.list({
      calendarId, q: query, timeMin, timeMax,
      singleEvents: true, orderBy: "startTime", maxResults: 25,
    });
    return textResult(res.data.items?.map(formatEvent) || []);
  });

  server.tool("calendar_respond_to_event", "Respond to a calendar event invitation", {
    calendarId: z.string(),
    eventId: z.string(),
    responseStatus: z.enum(["accepted", "declined", "tentative"]),
    sendUpdates: z.enum(["all", "externalOnly", "none"]).optional(),
  }, async ({ calendarId, eventId, responseStatus, sendUpdates }) => {
    const cal = api();
    const existing = await cal.events.get({ calendarId, eventId });
    const profile = await google.oauth2({ version: "v2", auth: ctx.auth }).userinfo.get();
    const myEmail = profile.data.email;

    const attendees = existing.data.attendees || [];
    const me = attendees.find((a) => a.email === myEmail || a.self);
    if (me) me.responseStatus = responseStatus;

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
    const res = await api().freebusy.query({
      requestBody: {
        timeMin, timeMax, timeZone,
        items: calendarIds.map((id) => ({ id })),
      },
    });
    return textResult(res.data.calendars);
  });

  server.tool("calendar_list_calendars", "List all calendars", {}, async () => {
    const res = await api().calendarList.list();
    return textResult(res.data.items?.map((c) => ({
      id: c.id, summary: c.summary, primary: c.primary, accessRole: c.accessRole, timeZone: c.timeZone,
    })) || []);
  });

  server.tool("calendar_list_colors", "List available event and calendar colors", {}, async () => {
    const res = await api().colors.get();
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
