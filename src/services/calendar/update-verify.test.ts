import { afterEach, describe, expect, it, vi } from "vitest";
import { z } from "zod";
import { google, calendar_v3 } from "googleapis";
import { formatEvent, registerCalendarTools } from "./index.js";

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
    expect(text).toMatch(/This calendar's copy of the event is not the organizer's copy/);
    expect(text).not.toMatch(/authenticated account/);
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

describe("calendar_update_event skips comparisons it cannot decide", () => {
  const ok = (r: { result: { isError?: boolean }; text: string }) => {
    expect(r.result.isError).toBeUndefined();
    expect(JSON.parse(r.text).id).toBeDefined();
  };
  const at = (dateTime: string, timeZone?: string) => ({ ...guestInstance, start: { dateTime, timeZone } });

  it("a wall-clock time inside a spring-forward gap", async () => {
    ok(await update({ start: "2026-03-08T02:30:00", timeZone: "America/New_York" }, at("2026-03-08T03:30:00-04:00", "America/New_York")));
  });

  it("an ambiguous fall-back hour, whichever occurrence Google picked", async () => {
    ok(await update({ start: "2026-11-01T01:30:00", timeZone: "America/New_York" }, at("2026-11-01T01:30:00-05:00", "America/New_York")));
    ok(await update({ start: "2026-11-01T01:30:00", timeZone: "America/New_York" }, at("2026-11-01T01:30:00-04:00", "America/New_York")));
  });

  it("an undecidable boundary is listed as not checked, never as applied", async () => {
    const { result, text } = await update(
      { start: "2026-03-08T02:30:00", timeZone: "America/New_York", location: "Room 9" },
      at("2026-03-08T03:30:00-04:00", "America/New_York"),
    );
    expect(result.isError).toBe(true);
    expect(text).toContain("did not apply: location.");
    expect(text).toContain("checked and applied: none.");
    expect(text).toContain("not checked: start.");
  });

  it("offset forms other than Z or ±HH:MM", async () => {
    ok(await update({ start: "2026-10-06T19:00:00+0000", timeZone: "America/New_York" }, at("2026-10-06T15:00:00-04:00", "America/New_York")));
    ok(await update({ start: "2026-10-07T00:00:00+05", timeZone: "America/New_York" }, at("2026-10-06T15:00:00-04:00", "America/New_York")));
  });

  it("a local time with no timeZone sent and none on the existing event", async () => {
    ok(await update({ start: "2026-10-06T15:00:00" }, at("2026-10-06T15:00:00-07:00", "UTC"), { start: {}, end: {} }));
  });

  it("description is not compared, since Google may sanitize it", async () => {
    ok(await update({ description: "<b>agenda</b>" }, { ...guestInstance, description: "<b>agenda</b><br>" }));
  });

  it("summary and location compare after trimming and CRLF normalization", async () => {
    ok(await update(
      { summary: "  Renamed\r\nsync \n", location: "\tRoom 9 " },
      { ...guestInstance, summary: "Renamed\nsync", location: "Room 9" },
    ));
  });
});

describe("calendar_update_event mismatch report wording", () => {
  it("names applied, unapplied and unchecked fields, with no hint when the organizer is absent", async () => {
    const { organizer, ...noOrganizer } = guestInstance;
    void organizer;
    const { result, text } = await update(
      { summary: "Renamed sync", start: "2026-10-06T15:00:00-04:00", description: "d", reminders: { useDefault: true }, recurrence: ["RRULE:FREQ=WEEKLY"] },
      noOrganizer,
    );
    expect(result.isError).toBe(true);
    expect(text).toContain("did not apply: start.");
    expect(text).toContain("checked and applied: summary.");
    expect(text).toContain("not checked: reminders, recurrence, description.");
    expect(text).not.toMatch(/other requested fields applied/i);
    expect(text).not.toMatch(/not this event.s organizer/);
  });
});

describe("calendar_update_event notes a guest-copy-only change on success", () => {
  const { recurringEventId, originalStartTime, ...single } = guestInstance;
  void recurringEventId; void originalStartTime;
  const moved = { ...single, id: "single1",
    start: { dateTime: "2026-10-06T15:00:00-04:00", timeZone: "America/New_York" },
    end: { dateTime: "2026-10-06T16:00:00-04:00", timeZone: "America/New_York" } };
  const input = { summary: "Renamed sync", start: "2026-10-06T15:00:00-04:00", end: "2026-10-06T16:00:00-04:00" };

  it("a guest whose shared-field change applied gets the note", async () => {
    const { result, text } = await update(input, moved);
    expect(result.isError).toBeUndefined();
    const body = JSON.parse(text);
    expect(body.note).toContain("This calendar's copy of the event is not the organizer's copy: any change to start, end, summary applies to this calendar's copy only; the organizer and other guests keep their own values.");
    expect(body.note).not.toMatch(/authenticated account|previous values|your copy/);
    expect(body.id).toBe("single1");
  });

  it("a guest changing only private fields gets no note", async () => {
    const { text } = await update({ colorId: "5" }, { ...moved, colorId: "5" });
    expect(JSON.parse(text).note).toBeUndefined();
  });

  it("the organizer's success payload is unchanged", async () => {
    const own = { ...moved, organizer: { email: "me@example.com", self: true } };
    const { text } = await update(input, own);
    expect(JSON.parse(text).note).toBeUndefined();
    expect(text).toBe(JSON.stringify(formatEvent(own)));
  });

  it("no organizer in the response, no note", async () => {
    const { organizer, ...anon } = moved;
    void organizer;
    const { text } = await update(input, anon);
    expect(JSON.parse(text).note).toBeUndefined();
  });

  it("the error path keeps its own hint and adds no note", async () => {
    const { text } = await update(input, guestInstance);
    expect(text).toMatch(/This calendar's copy of the event is not the organizer's copy/);
    expect(text).not.toMatch(/applies to this calendar's copy only/);
  });
});

describe("calendar_update_event round-two coverage", () => {
  it("an ambiguous fall-back hour in a zone at +12h or more is not flagged", async () => {
    const at = (dateTime: string, timeZone: string) => ({ ...guestInstance, start: { dateTime, timeZone } });
    for (const got of ["2026-04-05T02:30:00+13:00", "2026-04-05T02:30:00+12:00"]) {
      const { result } = await update({ start: "2026-04-05T02:30:00", timeZone: "Pacific/Auckland" }, at(got, "Pacific/Auckland"));
      expect(result.isError).toBeUndefined();
    }
    for (const got of ["2026-04-05T03:15:00+13:45", "2026-04-05T03:15:00+12:45"]) {
      const { result } = await update({ start: "2026-04-05T03:15:00", timeZone: "Pacific/Chatham" }, at(got, "Pacific/Chatham"));
      expect(result.isError).toBeUndefined();
    }
  });

  it("a recurrence change on a guest's copy carries the note", async () => {
    const { text } = await update({ recurrence: ["RRULE:FREQ=DAILY"] }, { ...guestInstance, recurrence: ["RRULE:FREQ=DAILY"] });
    expect(JSON.parse(text).note).toContain("any change to recurrence applies to this calendar's copy only");
  });

  it("the tool description names every checked, unchecked and shared field", async () => {
    let description = "";
    registerCalendarTools({ tool: (n: string, d: string) => { if (n === "calendar_update_event") description = d; } } as never, { auth: {} as never });
    for (const f of ["start, end, summary, location, colorId, visibility, transparency, guestsCanModify, guestsCanInviteOthers, guestsCanSeeOtherGuests",
      "attendees, reminders, recurrence, description", "start, end, summary, location, description, recurrence"]) {
      expect(description).toContain(f);
    }
  });
});
