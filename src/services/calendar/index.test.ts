import { afterEach, describe, expect, it, vi } from "vitest";
import { z } from "zod";
import { google, calendar_v3 } from "googleapis";
import { formatEventForList, registerCalendarTools } from "./index.js";

describe("formatEventForList", () => {
  const baseEvent: calendar_v3.Schema$Event = {
    id: "evt1",
    summary: "Team sync",
    description: "Weekly sync",
    start: { dateTime: "2026-07-06T10:00:00Z" },
    end: { dateTime: "2026-07-06T11:00:00Z" },
    location: "Room 1",
    status: "confirmed",
    htmlLink: "https://calendar.google.com/event?eid=abc",
    attendees: [
      { email: "alice@example.com", displayName: "Alice", responseStatus: "accepted", organizer: true, self: true },
      { email: "bob@example.com", displayName: "Bob", responseStatus: "needsAction" },
    ],
    conferenceData: {
      conferenceId: "conf123",
      entryPoints: [{ entryPointType: "video", uri: "https://meet.google.com/abc-defg-hij" }],
    },
    recurrence: ["RRULE:FREQ=WEEKLY"],
    creator: { email: "alice@example.com" },
    organizer: { email: "alice@example.com" },
  };

  it("trims attendees to email + responseStatus only", () => {
    const result = formatEventForList(baseEvent, false);
    expect(result.attendees).toEqual([
      { email: "alice@example.com", responseStatus: "accepted" },
      { email: "bob@example.com", responseStatus: "needsAction" },
    ]);
  });

  it("drops conferenceData when includeDetails is false", () => {
    const result = formatEventForList(baseEvent, false);
    expect(result.conferenceData).toBeUndefined();
    expect("conferenceData" in JSON.parse(JSON.stringify(result))).toBe(false);
  });

  it("keeps full conferenceData when includeDetails is true", () => {
    const result = formatEventForList(baseEvent, true);
    expect(result.conferenceData).toEqual(baseEvent.conferenceData);
  });

  it("still trims attendees even when includeDetails is true", () => {
    const result = formatEventForList(baseEvent, true);
    expect(result.attendees).toEqual([
      { email: "alice@example.com", responseStatus: "accepted" },
      { email: "bob@example.com", responseStatus: "needsAction" },
    ]);
  });

  it("preserves the other formatEvent fields unchanged", () => {
    const result = formatEventForList(baseEvent, false);
    expect(result.id).toBe("evt1");
    expect(result.summary).toBe("Team sync");
    expect(result.start).toEqual(baseEvent.start);
    expect(result.end).toEqual(baseEvent.end);
    expect(result.recurrence).toEqual(["RRULE:FREQ=WEEKLY"]);
  });

  it("handles an event with no attendees", () => {
    const result = formatEventForList({ ...baseEvent, attendees: undefined }, false);
    expect(result.attendees).toBeUndefined();
  });
});

const ATTENDEE_KEYS = [
  "additionalGuests", "comment", "displayName", "email", "id",
  "optional", "organizer", "resource", "responseStatus", "self",
] as const;
// Compile-time: fails to typecheck if Schema$EventAttendee gains a key missing from ATTENDEE_KEYS.
type MissingKeys = Exclude<keyof calendar_v3.Schema$EventAttendee, (typeof ATTENDEE_KEYS)[number]>;
const _exhaustive: [MissingKeys] extends [never] ? true : never = true;
void _exhaustive;

afterEach(() => { vi.restoreAllMocks(); });

describe("calendar attendee schema (create/update)", () => {
  type Handler = (opts: Record<string, unknown>) => Promise<unknown>;
  const setup = async () => {
    const patch = vi.fn().mockResolvedValue({ data: { id: "evt1" } });
    const get = vi.fn().mockResolvedValue({ data: { etag: '"e1"', attendees: [] } });
    vi.spyOn(google, "calendar").mockReturnValue({ events: { patch, get } } as never);
    const tools = new Map<string, { shape: z.ZodRawShape; handler: Handler }>();
    const server = { tool: (name: string, _d: string, shape: z.ZodRawShape, handler: Handler) => { tools.set(name, { shape, handler }); } };
    registerCalendarTools(server as never, { auth: {} as never });
    return { patch, get, tools };
  };

  // Shape of calendar_get_event output (Google's raw attendee objects).
  const readAttendees = [
    { email: "a@example.com", displayName: "A", responseStatus: "accepted", self: true, organizer: true, id: "123" },
    { email: "b@example.com", optional: true, responseStatus: "needsAction", comment: "maybe", additionalGuests: 2 },
    { email: "room@example.com", resource: true, responseStatus: "accepted" },
  ];

  const forwarded = [
    { email: "a@example.com", displayName: "A" },
    { email: "b@example.com", optional: true },
    { email: "room@example.com", resource: true },
  ];

  it("update_event accepts get_event-shaped attendees and forwards only organizer-controlled fields", async () => {
    const { patch, tools } = await setup();
    const t = tools.get("calendar_update_event")!;
    const parsed = z.object(t.shape).parse({ calendarId: "primary", eventId: "evt1", attendees: readAttendees });
    await t.handler(parsed);
    expect(patch.mock.calls[0][0].requestBody.attendees).toEqual(forwarded);
  });

  it("update_event accepts optional on an attendee", async () => {
    const { patch, tools } = await setup();
    const t = tools.get("calendar_update_event")!;
    const parsed = z.object(t.shape).parse({ calendarId: "primary", eventId: "evt1", attendees: [{ email: "b@example.com", optional: true }] });
    await t.handler(parsed);
    expect(patch.mock.calls[0][0].requestBody.attendees).toEqual([{ email: "b@example.com", optional: true }]);
  });

  it("create_event and update_event use one identical attendee schema covering every EventAttendee key", async () => {
    const { tools } = await setup();
    const elementOf = (name: string) => {
      const field = tools.get(name)!.shape.attendees as z.ZodOptional<z.ZodArray<z.ZodObject<z.ZodRawShape>>>;
      return field.unwrap().element;
    };
    const c = elementOf("calendar_create_event");
    const u = elementOf("calendar_update_event");
    expect(c).toBe(u);
    expect(Object.keys(c.shape).sort()).toEqual([...ATTENDEE_KEYS].sort());
    expect(Object.keys(u.shape).sort()).toEqual([...ATTENDEE_KEYS].sort());
    // Google's `T | null` typings: round-trip-only fields accept null.
    expect(c.safeParse({ email: "a@example.com", responseStatus: null, comment: null, additionalGuests: null, id: null, self: null, organizer: null }).success).toBe(true);
  });
});

describe("calendar_create_event attendee forwarding", () => {
  it("forwards only organizer-controlled attendee fields", async () => {
    const insert = vi.fn().mockResolvedValue({ data: { id: "evt1" } });
    vi.spyOn(google, "calendar").mockReturnValue({ events: { insert }, calendars: { get: vi.fn().mockResolvedValue({ data: { timeZone: "UTC" } }) } } as never);
    const tools = new Map<string, { shape: z.ZodRawShape; handler: (o: Record<string, unknown>) => Promise<unknown> }>();
    registerCalendarTools({ tool: (n: string, _d: string, shape: z.ZodRawShape, handler: never) => { tools.set(n, { shape, handler }); } } as never, { auth: {} as never });
    const t = tools.get("calendar_create_event")!;
    const parsed = z.object(t.shape).parse({
      calendarId: "primary", summary: "x", start: "2026-07-06T10:00:00Z", end: "2026-07-06T11:00:00Z",
      attendees: [{ email: "b@example.com", optional: true, responseStatus: "accepted", comment: "c", self: true }],
    });
    await t.handler(parsed);
    expect(insert.mock.calls[0][0].requestBody.attendees).toEqual([{ email: "b@example.com", optional: true }]);
  });
});

describe("calendar_update_event RSVP preservation", () => {
  type Handler = (opts: Record<string, unknown>) => Promise<unknown>;
  const run = async (liveData: unknown, input: Record<string, unknown>) => {
    const patch = vi.fn().mockResolvedValue({ data: { id: "evt1" } });
    const get = vi.fn().mockResolvedValue({ data: { etag: '"e1"', ...(liveData as object) } });
    vi.spyOn(google, "calendar").mockReturnValue({ events: { patch, get } } as never);
    const tools = new Map<string, { shape: z.ZodRawShape; handler: Handler }>();
    registerCalendarTools({ tool: (n: string, _d: string, shape: z.ZodRawShape, handler: Handler) => { tools.set(n, { shape, handler }); } } as never, { auth: {} as never });
    const t = tools.get("calendar_update_event")!;
    await t.handler(z.object(t.shape).parse({ calendarId: "primary", eventId: "evt1", ...input }));
    return { patch, get };
  };

  it("carries live responseStatus/comment/additionalGuests for existing attendees, ignoring stale input; new attendees get none", async () => {
    const { patch } = await run(
      { attendees: [
        { email: "a@x.com", responseStatus: "accepted", comment: "yes", additionalGuests: 1 },
        { email: "b@x.com", responseStatus: "declined" },
      ] },
      { attendees: [
        { email: "a@x.com", responseStatus: "needsAction", comment: "stale", additionalGuests: 9 },
        { email: "b@x.com" },
        { email: "c@x.com", displayName: "C" },
      ] },
    );
    expect(patch.mock.calls[0][0].requestBody.attendees).toEqual([
      { email: "a@x.com", responseStatus: "accepted", comment: "yes", additionalGuests: 1 },
      { email: "b@x.com", responseStatus: "declined" },
      { email: "c@x.com", displayName: "C" },
    ]);
  });

  it("matches emails case-insensitively and keeps the caller's email spelling", async () => {
    const { patch } = await run(
      { attendees: [{ email: "Alice@X.com", responseStatus: "tentative" }] },
      { attendees: [{ email: "alice@x.COM" }] },
    );
    expect(patch.mock.calls[0][0].requestBody.attendees).toEqual([{ email: "alice@x.COM", responseStatus: "tentative" }]);
  });

  it("does not fetch the event when attendees is not provided", async () => {
    const { patch, get } = await run({}, { summary: "new title" });
    expect(get).not.toHaveBeenCalled();
    expect(patch.mock.calls[0][0].requestBody).toEqual({ summary: "new title" });
  });

  it("does at most one events.get when both attendees and the timezone lookup need the event", async () => {
    const { patch, get } = await run(
      { attendees: [{ email: "a@x.com", responseStatus: "accepted" }], start: { timeZone: "America/New_York" } },
      { attendees: [{ email: "a@x.com" }], start: "2026-07-06T10:00:00" },
    );
    expect(get).toHaveBeenCalledTimes(1);
    const body = patch.mock.calls[0][0].requestBody;
    expect(body.attendees).toEqual([{ email: "a@x.com", responseStatus: "accepted" }]);
    expect(body.start).toEqual({ dateTime: "2026-07-06T10:00:00", timeZone: "America/New_York" });
  });
});

type FakeAttendee = calendar_v3.Schema$EventAttendee;
type FakeStore = { attendees: FakeAttendee[]; version: number; bump: () => void };

// In-memory stand-in for one event on Google Calendar. Every write bumps the
// etag, a patch carrying If-Match is refused with 412 when the etag no longer
// matches, and an attendees array in a patch replaces the stored list whole,
// as Google does.
const fakeCalendar = (initial: FakeAttendee[], opts: { beforeEachPatch?: (store: FakeStore) => void; omitEtag?: boolean } = {}) => {
  const store: FakeStore = { attendees: structuredClone(initial), version: 1, bump: () => { store.version++; } };
  const etagOf = () => `"v${store.version}"`;
  const get = vi.fn(async (_params?: unknown) => ({ data: { id: "evt1", etag: opts.omitEtag ? undefined : etagOf(), attendees: structuredClone(store.attendees) } }));
  const patch = vi.fn(async (params: { requestBody: { attendees?: FakeAttendee[] } }, options?: { headers?: Record<string, string> }) => {
    opts.beforeEachPatch?.(store);
    const ifMatch = options?.headers?.["If-Match"];
    if (ifMatch !== undefined && ifMatch !== etagOf()) {
      throw Object.assign(new Error("Precondition Failed"), { response: { status: 412 } });
    }
    if (params.requestBody.attendees) store.attendees = structuredClone(params.requestBody.attendees);
    store.bump();
    return { data: { id: "evt1", etag: etagOf(), attendees: structuredClone(store.attendees) } };
  });
  return { store, get, patch };
};

// Registers the calendar tools from a freshly loaded module wired to the fake,
// so the module's process-lifetime primary-email cache never carries over
// from another test, and runs one tool.
const runTool = async (fake: ReturnType<typeof fakeCalendar>, tool: string, input: Record<string, unknown>, primaryEmail?: string) => {
  type Handler = (opts: Record<string, unknown>) => Promise<unknown>;
  vi.resetModules();
  const { google: g } = await import("googleapis");
  const mod = await import("./index.js");
  const calendarListGet = vi.fn(async () => ({ data: { id: primaryEmail } }));
  vi.spyOn(g, "calendar").mockReturnValue({ events: { get: fake.get, patch: fake.patch }, calendarList: { get: calendarListGet } } as never);
  const tools = new Map<string, { shape: z.ZodRawShape; handler: Handler }>();
  mod.registerCalendarTools({ tool: (n: string, _d: string, shape: z.ZodRawShape, handler: Handler) => { tools.set(n, { shape, handler }); } } as never, { auth: {} as never });
  const t = tools.get(tool)!;
  return t.handler(z.object(t.shape).parse({ calendarId: "primary", eventId: "evt1", ...input }));
};

describe("calendar_respond_to_event", () => {
  type Attendee = FakeAttendee;
  const respond = (fake: ReturnType<typeof fakeCalendar>, input: Record<string, unknown>, primaryEmail?: string) =>
    runTool(fake, "calendar_respond_to_event", input, primaryEmail);

  const initial: Attendee[] = [
    { email: "me@example.com", self: true, responseStatus: "needsAction", comment: "running late", additionalGuests: 2 },
    { email: "other@example.com", responseStatus: "needsAction" },
    { email: "third@example.com", displayName: "Third", optional: true, responseStatus: "tentative", comment: "maybe", additionalGuests: 1 },
  ];

  it("keeps an RSVP another attendee makes between the read and the write", async () => {
    let raced = false;
    const fake = fakeCalendar(initial, {
      beforeEachPatch: (s) => {
        if (raced) return;
        raced = true;
        s.attendees = s.attendees.map((a) => (a.email === "other@example.com" ? { ...a, responseStatus: "accepted", comment: "see you there" } : a));
        s.bump();
      },
    });
    await respond(fake, { responseStatus: "declined" });
    expect(fake.store.attendees).toEqual([
      { ...initial[0], responseStatus: "declined" },
      { email: "other@example.com", responseStatus: "accepted", comment: "see you there" },
      initial[2],
    ]);
  });

  it("writes once, conditional on the read's etag, changing only the caller's responseStatus", async () => {
    const fake = fakeCalendar(initial);
    await respond(fake, { responseStatus: "accepted", sendUpdates: "none" });
    expect(fake.patch).toHaveBeenCalledTimes(1);
    expect(fake.patch.mock.calls[0][1]).toEqual({ headers: { "If-Match": '"v1"' } });
    expect(fake.patch.mock.calls[0][0]).toMatchObject({ calendarId: "primary", eventId: "evt1", sendUpdates: "none" });
    expect(fake.store.attendees).toEqual([{ ...initial[0], responseStatus: "accepted" }, initial[1], initial[2]]);
  });

  it("finds the caller by primary-calendar email when no attendee is flagged self", async () => {
    const noSelf: Attendee[] = [{ email: "Someone@Example.com", responseStatus: "needsAction" }, { email: "other@example.com", responseStatus: "accepted" }];
    const fake = fakeCalendar(noSelf);
    await respond(fake, { responseStatus: "tentative" }, "someone@example.com");
    expect(fake.store.attendees).toEqual([{ email: "Someone@Example.com", responseStatus: "tentative" }, noSelf[1]]);
  });

  it("gives up after 3 attempts with an error, having written nothing, when the event changes every time", async () => {
    const fake = fakeCalendar(initial, { beforeEachPatch: (s) => s.bump() });
    const err = await respond(fake, { responseStatus: "declined" }).catch((e: Error) => e);
    expect(err).toBeInstanceOf(Error);
    expect((err as Error).message).toMatch(/nothing was written/);
    expect(((err as Error).cause as { response?: { status?: number } }).response?.status).toBe(412);
    expect(fake.get).toHaveBeenCalledTimes(3);
    expect(fake.patch).toHaveBeenCalledTimes(3);
    expect(fake.store.attendees).toEqual(initial);
  });

  it("does not retry errors other than a precondition failure", async () => {
    const fake = fakeCalendar(initial, {
      beforeEachPatch: () => { throw Object.assign(new Error("Forbidden"), { response: { status: 403 } }); },
    });
    await expect(respond(fake, { responseStatus: "declined" })).rejects.toThrow(/Forbidden/);
    expect(fake.patch).toHaveBeenCalledTimes(1);
  });

  it("refuses to write when the read carries no etag", async () => {
    const fake = fakeCalendar(initial, { omitEtag: true });
    await expect(respond(fake, { responseStatus: "declined" })).rejects.toThrow(/without an etag/);
    expect(fake.patch).not.toHaveBeenCalled();
  });

  it("refuses, without writing, when the caller is not an attendee", async () => {
    const fake = fakeCalendar([{ email: "other@example.com", responseStatus: "accepted" }]);
    await expect(respond(fake, { responseStatus: "accepted" }, "me@example.com")).rejects.toThrow(/not an attendee/);
    expect(fake.patch).not.toHaveBeenCalled();
  });
});

describe("calendar_update_event attendee writes", () => {
  const update = (fake: ReturnType<typeof fakeCalendar>, input: Record<string, unknown>) => runTool(fake, "calendar_update_event", input);
  const live: FakeAttendee[] = [
    { email: "a@example.com", responseStatus: "accepted", comment: "yes" },
    { email: "b@example.com", responseStatus: "needsAction" },
  ];
  const input = { attendees: [{ email: "a@example.com" }, { email: "b@example.com" }, { email: "c@example.com", displayName: "C" }] };

  it("keeps an RSVP an attendee makes between the read and the write", async () => {
    let raced = false;
    const fake = fakeCalendar(live, {
      beforeEachPatch: (s) => {
        if (raced) return;
        raced = true;
        s.attendees = s.attendees.map((a) => (a.email === "b@example.com" ? { ...a, responseStatus: "accepted", comment: "on my way" } : a));
        s.bump();
      },
    });
    await update(fake, input);
    expect(fake.store.attendees).toEqual([
      { email: "a@example.com", responseStatus: "accepted", comment: "yes" },
      { email: "b@example.com", responseStatus: "accepted", comment: "on my way" },
      { email: "c@example.com", displayName: "C" },
    ]);
  });

  it("writes once, conditional on the etag its read asked for", async () => {
    const fake = fakeCalendar(live);
    await update(fake, input);
    expect(fake.get).toHaveBeenCalledTimes(1);
    expect((fake.get.mock.calls[0][0] as { fields: string }).fields.split(",")).toContain("etag");
    expect(fake.patch).toHaveBeenCalledTimes(1);
    expect(fake.patch.mock.calls[0][1]).toEqual({ headers: { "If-Match": '"v1"' } });
  });

  it("gives up after 3 attempts with an error, having written nothing, when the event changes every time", async () => {
    const fake = fakeCalendar(live, { beforeEachPatch: (s) => s.bump() });
    const err = await update(fake, input).catch((e: Error) => e);
    expect((err as Error).message).toMatch(/nothing was written/);
    expect(fake.patch).toHaveBeenCalledTimes(3);
    expect(fake.store.attendees).toEqual(live);
  });

  it("refuses to write attendees when the read carries no etag", async () => {
    const fake = fakeCalendar(live, { omitEtag: true });
    await expect(update(fake, input)).rejects.toThrow(/without an etag/);
    expect(fake.patch).not.toHaveBeenCalled();
  });

  it("leaves an update without attendees unconditional and unread", async () => {
    const fake = fakeCalendar(live);
    await update(fake, { summary: "new title" });
    expect(fake.get).not.toHaveBeenCalled();
    expect(fake.patch.mock.calls[0][1]).toBeUndefined();
  });
});
