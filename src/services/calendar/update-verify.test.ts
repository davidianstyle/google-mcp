import { afterEach, describe, expect, it, vi } from "vitest";
import { z } from "zod";
import { google, calendar_v3 } from "googleapis";
import { registerCalendarTools } from "./index.js";

type Handler = (opts: Record<string, unknown>) => Promise<{ content: Array<{ text: string }>; isError?: boolean }>;

afterEach(() => { vi.restoreAllMocks(); });

// Runs calendar_update_event with `patched` as the events.patch response and
// `live` as the events.get response (the existing-timezone lookup).
async function update(input: Record<string, unknown>, patched: calendar_v3.Schema$Event, live: calendar_v3.Schema$Event = {}) {
  const patch = vi.fn().mockResolvedValue({ data: patched });
  const get = vi.fn().mockResolvedValue({ data: live });
  vi.spyOn(google, "calendar").mockReturnValue({ events: { patch, get } } as never);
  const tools = new Map<string, { shape: z.ZodRawShape; handler: Handler }>();
  registerCalendarTools({ tool: (n: string, _d: string, shape: z.ZodRawShape, handler: Handler) => { tools.set(n, { shape, handler }); } } as never, { auth: {} as never });
  const t = tools.get("calendar_update_event")!;
  const result = await t.handler(z.object(t.shape).parse({ calendarId: "primary", eventId: "evt1", ...input }));
  return { result, patch, text: result.content[0].text };
}

// An instance of a recurring series the authenticated user is a guest on, as
// the API returns it from events.patch: offset-qualified dateTimes rendered
// in the calendar's zone, the event's own zone in timeZone.
const guestInstance: calendar_v3.Schema$Event = {
  kind: "calendar#event",
  id: "series1_20261006T140000Z",
  recurringEventId: "series1",
  originalStartTime: { dateTime: "2026-10-06T10:00:00-04:00", timeZone: "America/New_York" },
  status: "confirmed",
  summary: "Renamed sync",
  start: { dateTime: "2026-10-06T10:00:00-04:00", timeZone: "America/New_York" },
  end: { dateTime: "2026-10-06T11:00:00-04:00", timeZone: "America/New_York" },
  organizer: { email: "organizer@example.com" },
  creator: { email: "organizer@example.com" },
  attendees: [
    { email: "organizer@example.com", organizer: true, responseStatus: "accepted" },
    { email: "me@example.com", self: true, responseStatus: "accepted" },
  ],
  reminders: { useDefault: true },
};

describe("calendar_update_event reports fields the API did not apply", () => {
  it("flags start/end that came back unchanged while the summary applied", async () => {
    const { result, text } = await update(
      { summary: "Renamed sync", start: "2026-10-06T15:00:00-04:00", end: "2026-10-06T16:00:00-04:00" },
      guestInstance,
    );
    expect(result.isError).toBe(true);
    expect(text).toMatch(/did not apply: start, end\b/);
    expect(text).not.toMatch(/did not apply:[^.]*summary/);
    expect(text).toMatch(/not this event's organizer/);
    expect(text).toContain("2026-10-06T10:00:00-04:00");
  });

  it("flags a single unapplied boundary and a scalar field", async () => {
    const { result, text } = await update(
      { end: "2026-10-06T12:00:00-04:00", location: "Room 9" },
      guestInstance,
    );
    expect(result.isError).toBe(true);
    expect(text).toMatch(/did not apply: end, location\./);
  });

  it("flags an all-day change that left the event timed", async () => {
    const { result, text } = await update({ start: "2026-10-06", end: "2026-10-07" }, guestInstance);
    expect(result.isError).toBe(true);
    expect(text).toMatch(/did not apply: start, end\b/);
  });

  it("flags a boolean that came back at its old non-default value", async () => {
    const { result, text } = await update(
      { guestsCanModify: true },
      { ...guestInstance, organizer: { email: "me@example.com", self: true } },
    );
    expect(result.isError).toBe(true);
    expect(text).toMatch(/did not apply: guestsCanModify\./);
    expect(text).not.toMatch(/organizer;/);
  });
});

describe("calendar_update_event does not flag writes that applied", () => {
  const ok = (r: { result: { isError?: boolean }; text: string }) => {
    expect(r.result.isError).toBeUndefined();
    expect(JSON.parse(r.text).id).toBeDefined();
  };

  it("offset-less local time + explicit timeZone, returned rendered in another zone's offset", async () => {
    ok(await update(
      { start: "2026-10-06T10:00:00", end: "2026-10-06T11:00:00", timeZone: "America/Los_Angeles" },
      { ...guestInstance,
        start: { dateTime: "2026-10-06T13:00:00-04:00", timeZone: "America/Los_Angeles" },
        end: { dateTime: "2026-10-06T14:00:00-04:00", timeZone: "America/Los_Angeles" } },
    ));
  });

  it("offset-less local time resolved through the event's existing timezone", async () => {
    const r = await update(
      { start: "2026-10-06T15:00:00", end: "2026-10-06T16:00:00" },
      { ...guestInstance,
        start: { dateTime: "2026-10-06T15:00:00-04:00", timeZone: "America/New_York" },
        end: { dateTime: "2026-10-06T16:00:00-04:00", timeZone: "America/New_York" } },
      { start: { timeZone: "America/New_York" }, end: { timeZone: "America/New_York" } },
    );
    ok(r);
    expect(r.patch.mock.calls[0][0].requestBody.start).toEqual({ dateTime: "2026-10-06T15:00:00", timeZone: "America/New_York" });
  });

  it("UTC 'Z' input against an offset-rendered response, with fractional seconds", async () => {
    ok(await update(
      { start: "2026-10-06T19:00:00.000Z", end: "2026-10-06T20:00:00Z" },
      { ...guestInstance,
        start: { dateTime: "2026-10-06T15:00:00-04:00", timeZone: "America/New_York" },
        end: { dateTime: "2026-10-06T16:00:00-04:00", timeZone: "America/New_York" } },
    ));
  });

  it("local time just after a DST change resolves with the offset in force at that hour", async () => {
    ok(await update(
      { start: "2026-11-01T03:00:00", end: "2026-11-01T04:00:00", timeZone: "America/New_York" },
      { ...guestInstance,
        start: { dateTime: "2026-11-01T08:00:00Z", timeZone: "America/New_York" },
        end: { dateTime: "2026-11-01T09:00:00Z", timeZone: "America/New_York" } },
    ));
  });

  it("all-day dates that applied", async () => {
    ok(await update(
      { start: "2026-10-06", end: "2026-10-07" },
      { ...guestInstance, start: { date: "2026-10-06" }, end: { date: "2026-10-07" } },
    ));
  });

  it("recurring instance moved off its original slot (originalStartTime keeps the old time)", async () => {
    ok(await update(
      { start: "2026-10-06T15:00:00-04:00", end: "2026-10-06T16:00:00-04:00" },
      { ...guestInstance,
        start: { dateTime: "2026-10-06T15:00:00-04:00", timeZone: "America/New_York" },
        end: { dateTime: "2026-10-06T16:00:00-04:00", timeZone: "America/New_York" } },
    ));
  });

  it("default values the API omits from the response", async () => {
    const { visibility, transparency, guestsCanModify, guestsCanInviteOthers, guestsCanSeeOtherGuests, ...bare } =
      { ...guestInstance, visibility: "private", transparency: "transparent", guestsCanModify: true, guestsCanInviteOthers: false, guestsCanSeeOtherGuests: false };
    void visibility; void transparency; void guestsCanModify; void guestsCanInviteOthers; void guestsCanSeeOtherGuests;
    ok(await update(
      { visibility: "default", transparency: "opaque", guestsCanModify: false, guestsCanInviteOthers: true, guestsCanSeeOtherGuests: true, summary: "", description: "", location: "", colorId: "" },
      { ...bare, summary: undefined },
    ));
  });

  it("an unresolvable timezone is not treated as a mismatch", async () => {
    ok(await update(
      { start: "2026-10-06T15:00:00", timeZone: "Not/AZone" },
      { ...guestInstance, start: { dateTime: "2026-10-06T15:00:00-04:00", timeZone: "Not/AZone" } },
    ));
  });

  it("fields that were not requested are not checked", async () => {
    ok(await update({ summary: "Renamed sync" }, guestInstance));
  });
});
