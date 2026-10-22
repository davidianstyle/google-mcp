import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from "vitest";
import { z } from "zod";
import { OAuth2Client } from "google-auth-library";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer } from "../server.js";
import { assertToolsStrict } from "./errors.js";
import { deepStrict, newStrictifyStats } from "./strict-params.js";

function makeCtx() {
  return { auth: new OAuth2Client(), accountSlug: "test-slug" };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function schemaOf(server: McpServer, name: string): any {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return (server as any)._registeredTools[name].inputSchema;
}

describe("strict tool parameters (registration choke point)", () => {
  // Independent of the shell: these tests assume the default (strict) mode.
  let server: McpServer;
  beforeAll(() => {
    vi.stubEnv("GOOGLE_MCP_STRICT_PARAMS", undefined);
    server = createServer(makeCtx());
  });
  afterAll(() => {
    vi.unstubAllEnvs();
  });
  const event = { calendarId: "primary", summary: "s", start: "2026-10-01T10:00:00", end: "2026-10-01T11:00:00" };

  it("makes every registered tool strict", () => {
    const { strict, passthrough, noArgs } = assertToolsStrict(server);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect(strict + passthrough + noArgs).toBe(Object.keys((server as any)._registeredTools).length);
    expect(strict).toBeGreaterThan(100);
    expect(passthrough).toBe(0);
  });

  it("rejects a misspelled top-level parameter", () => {
    const r = schemaOf(server, "calendar_create_event").safeParse({ ...event, timezone: "UTC" });
    expect(r.success).toBe(false);
    expect(JSON.stringify(r.error.issues)).toContain("timezone");
  });

  it("rejects a misspelled nested parameter, including inside array items", () => {
    const s = schemaOf(server, "calendar_create_event");
    expect(s.safeParse({ ...event, reminders: { useDefault: false, override: [] } }).success).toBe(false);
    expect(s.safeParse({ ...event, attendees: [{ email: "a@b.c", optinal: true }] }).success).toBe(false);
  });

  it("accepts valid parameters unchanged, defaults still applied", () => {
    const r = schemaOf(server, "gmail_draft_email").safeParse({ to: ["a@b.c"], subject: "s", body: "b" });
    expect(r.success).toBe(true);
    expect(r.data.mimeType).toBe("text/plain");
  });

  it("accepts and ignores extra keys on a zero-parameter tool", () => {
    const schema = schemaOf(server, "gmail_list_labels");
    expect(schema.safeParse({}).success).toBe(true);
    const r = schema.safeParse({ random_string: "x" });
    expect(r.success).toBe(true);
    expect(r.data).toEqual({});
    expect(() => assertToolsStrict(server)).not.toThrow();
  });

  it("assertToolsStrict counts a zero-property object as a no-arg tool", () => {
    const s = createServer(makeCtx(), { families: ["meet"] });
    s.registerTool("noargs-bypass", { inputSchema: {} }, async () => ({ content: [] }));
    expect(() => assertToolsStrict(s)).not.toThrow();
  });

  it("assertToolsStrict accepts an explicit top-level passthrough as an opt-out", () => {
    const s = createServer(makeCtx(), { families: ["meet"] });
    s.tool("loose", { a: z.string() }, async () => ({ content: [] }));
    const before = assertToolsStrict(s);
    // simulate an explicit catchall schema registered by hand
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (s as any)._registeredTools["loose"].inputSchema = z.looseObject({ a: z.string() });
    const after = assertToolsStrict(s);
    // counted as passthrough, not strict
    expect(after.passthrough).toBe(before.passthrough + 1);
    expect(after.strict).toBe(before.strict - 1);
  });

  it("assertToolsStrict names the missing private SDK internal", () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect(() => assertToolsStrict({} as any)).toThrow(/_registeredTools/);
  });

  it("registers a tool with a transformed parameter and still rejects unknown nested keys", () => {
    const s = createServer(makeCtx(), { families: ["meet"] });
    s.tool(
      "pipe-ok",
      { opts: z.object({ n: z.string() }).transform((o) => ({ n: o.n.toUpperCase() })) },
      async () => ({ content: [] })
    );
    const schema = schemaOf(s, "pipe-ok");
    expect(schema.safeParse({ opts: { n: "a" } }).data).toEqual({ opts: { n: "A" } });
    expect(schema.safeParse({ opts: { n: "a", extra: 1 } }).success).toBe(false);
    expect(() => assertToolsStrict(s)).not.toThrow();
  });

  it("leaves no loosely-registered tool behind when strictification throws", () => {
    const s = createServer(makeCtx(), { families: ["meet"] });
    expect(() =>
      s.tool("lazy-gap", { o: z.lazy(() => z.object({ a: z.string() })) }, async () => ({ content: [] }))
    ).toThrow(/lazy-gap/);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect(Object.keys((s as any)._registeredTools)).not.toContain("lazy-gap");
    expect(() => assertToolsStrict(s)).not.toThrow();
  });

  it("fails the startup assertion when a tool bypasses the choke point", () => {
    const s = createServer(makeCtx(), { families: ["meet"] });
    s.registerTool("bypass", { inputSchema: { a: z.string() } }, async () => ({ content: [] }));
    expect(() => assertToolsStrict(s)).toThrow(/bypass/);
  });

  it("reaches objects inside a union and inside a record", () => {
    const s = createServer(makeCtx(), { families: ["meet"] });
    s.tool(
      "union-record",
      {
        u: z.union([z.object({ a: z.string() }), z.object({ b: z.string() })]).optional(),
        m: z.record(z.string(), z.object({ a: z.string() })).optional(),
      },
      async () => ({ content: [] })
    );
    const schema = schemaOf(s, "union-record");
    expect(schema.safeParse({ u: { a: "x" }, m: { k: { a: "x" } } }).success).toBe(true);
    expect(schema.safeParse({ u: { a: "x", extra: 1 } }).success).toBe(false);
    expect(schema.safeParse({ m: { k: { a: "x", extra: 1 } } }).success).toBe(false);
  });

  it("throws at registration when an object is nested under z.preprocess", () => {
    const s = createServer(makeCtx(), { families: ["meet"] });
    expect(() =>
      s.tool(
        "preprocess-gap",
        { opts: z.preprocess((v) => v, z.object({ timeZone: z.string().optional() })) },
        async () => ({ content: [] })
      )
    ).toThrow(/preprocess-gap/);
  });
});

describe("deepStrict", () => {
  it("preserves descriptions, checks, and function defaults on rebuilt wrappers", () => {
    let n = 0;
    const src = z
      .array(z.object({ a: z.string() }))
      .min(1)
      .default(() => [{ a: `v${++n}` }])
      .describe("the list");
    const out = deepStrict(src, newStrictifyStats());
    expect(out.description).toBe("the list");
    expect(out.safeParse([]).success).toBe(false); // .min(1) survived
    expect(out.parse(undefined)).toEqual([{ a: "v1" }]);
    expect(out.parse(undefined)).toEqual([{ a: "v2" }]); // default still re-evaluated
    expect(out.safeParse([{ a: "x", b: 1 }]).success).toBe(false);
  });

  it("leaves explicit passthrough objects alone", () => {
    const stats = newStrictifyStats();
    const out = deepStrict(z.looseObject({ a: z.string() }), stats);
    expect(out.safeParse({ a: "x", extra: 1 }).success).toBe(true);
    expect(stats.explicitCatchall).toBe(1);
  });
});

describe("deepStrict edge cases", () => {
  it("keeps walking below an object that is already .strict()", () => {
    const out = deepStrict(z.object({ o: z.strictObject({ n: z.object({ x: z.string() }) }) }), newStrictifyStats());
    expect(out.safeParse({ o: { n: { x: "a" } } }).success).toBe(true);
    expect(out.safeParse({ o: { n: { x: "a", typo: 1 } } }).success).toBe(false);
  });
  it("keeps an object's own error map for issues other than unrecognized keys", () => {
    const own = z.object(
      { a: z.string() },
      { error: (iss) => (iss.code === "invalid_type" ? "o must be an object" : undefined) }
    );
    const out = deepStrict(z.object({ o: own }), newStrictifyStats());
    const r = out.safeParse({ o: "not-an-object" });
    expect(r.success).toBe(false);
    expect(r.error.issues[0].message).toBe("o must be an object");
    expect(out.safeParse({ o: { a: "x", b: 1 } }).error.issues[0].message).toContain("Valid keys: a");
  });
  it("refuses .catch() rather than turning a typo into a silent fallback", () => {
    expect(() =>
      deepStrict(z.object({ o: z.object({ a: z.string() }).catch({ a: "DEFAULT" }) }), newStrictifyStats())
    ).toThrow(/catch/);
  });
});

describe("GOOGLE_MCP_STRICT_PARAMS runtime opt-out", () => {
  const open: Array<{ close(): Promise<void> }> = [];
  afterEach(async () => {
    for (const c of open.splice(0)) await c.close();
    vi.unstubAllEnvs();
    vi.restoreAllMocks();
  });

  // Builds the server exactly as startup does (env read inside createServer),
  // registers an argument-echoing tool through the same patched `server.tool`
  // every service uses, and connects a real MCP client to it.
  async function connect() {
    const s = createServer(makeCtx(), { families: ["calendar"] });
    s.tool("echo_args", { timeZone: z.string().optional() }, async (args) => ({
      content: [{ type: "text", text: JSON.stringify(args) }],
    }));
    s.tool(
      "nested_args",
      {
        range: z.object({ start: z.string(), end: z.string() }).optional(),
        none: z.object({}).optional(),
        authored: z
          .strictObject(
            { a: z.string() },
            { error: (i) => (i.code === "unrecognized_keys" ? "only 'a' is allowed here" : undefined) }
          )
          .optional(),
      },
      async (args) => ({ content: [{ type: "text", text: JSON.stringify(args) }] })
    );
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await s.connect(serverSide);
    const client = new Client({ name: "test", version: "0" });
    await client.connect(clientSide);
    open.push(client, s);
    return client;
  }
  const stray = { timeZone: "UTC", tmiezone: "UTC" };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const text = (r: any) => r.content[0].text as string;

  it("unset: a stray key is rejected with an input validation error", async () => {
    vi.stubEnv("GOOGLE_MCP_STRICT_PARAMS", undefined);
    const errSpy = vi.spyOn(console, "error").mockImplementation(() => {});
    const client = await connect();
    for (const name of ["echo_args", "calendar_get_current_time"]) {
      const r = await client.callTool({ name, arguments: stray });
      expect(r.isError).toBe(true);
      expect(text(r)).toMatch(/-32602/);
      expect(text(r)).toContain("tmiezone");
    }
    expect(errSpy).not.toHaveBeenCalledWith(expect.stringContaining("GOOGLE_MCP_STRICT_PARAMS"));
  });

  for (const value of ["1", "false", ""]) {
    it(`=${JSON.stringify(value)}: strict stays on through the request path`, async () => {
      vi.stubEnv("GOOGLE_MCP_STRICT_PARAMS", value);
      const errSpy = vi.spyOn(console, "error").mockImplementation(() => {});
      const client = await connect();
      const r = await client.callTool({ name: "calendar_get_current_time", arguments: stray });
      expect(r.isError).toBe(true);
      expect(text(r)).toContain('Unrecognized key: "tmiezone"');
      const lines = errSpy.mock.calls.filter((c) => String(c[0]).includes("GOOGLE_MCP_STRICT_PARAMS"));
      if (value === "") {
        expect(lines).toHaveLength(0);
      } else {
        expect(lines).toHaveLength(1);
        expect(String(lines[0][0])).toContain("not recognized");
      }
    });
  }

  it("unset: the rejection names the valid keys at the level of the stray key", async () => {
    vi.stubEnv("GOOGLE_MCP_STRICT_PARAMS", undefined);
    const client = await connect();
    const top = await client.callTool({ name: "calendar_get_current_time", arguments: stray });
    expect(text(top)).toMatch(/^MCP error -32602: /);
    expect(text(top)).toContain('Unrecognized key: "tmiezone". Valid keys: timeZone');
    const two = await client.callTool({ name: "calendar_get_current_time", arguments: { tz: "UTC", zone: "UTC" } });
    expect(text(two)).toContain('Unrecognized keys: "tz", "zone". Valid keys: timeZone');
    const nested = await client.callTool({
      name: "nested_args",
      arguments: { range: { start: "a", end: "b", ennd: "c" } },
    });
    expect(nested.isError).toBe(true);
    expect(text(nested)).toContain('Unrecognized key: "ennd". Valid keys: start, end at range');
    const empty = await client.callTool({ name: "nested_args", arguments: { none: { x: 1 } } });
    expect(text(empty)).toContain('Unrecognized key: "x". Valid keys: (none) at none');
  });

  it("unset: a schema's own unrecognized-key message wins over the valid-keys hint", async () => {
    vi.stubEnv("GOOGLE_MCP_STRICT_PARAMS", undefined);
    const client = await connect();
    const authored = await client.callTool({ name: "nested_args", arguments: { authored: { a: "x", b: 1 } } });
    expect(authored.isError).toBe(true);
    expect(text(authored)).toContain("only 'a' is allowed here at authored");
    expect(text(authored)).not.toContain("Valid keys");
    const plain = await client.callTool({ name: "nested_args", arguments: { range: { start: "a", end: "b", x: 1 } } });
    expect(text(plain)).toContain('Unrecognized key: "x". Valid keys: start, end at range');
  });

  it("=0: the same call succeeds and the stray key is dropped; one stderr line names the variable", async () => {
    vi.stubEnv("GOOGLE_MCP_STRICT_PARAMS", "0");
    const errSpy = vi.spyOn(console, "error").mockImplementation(() => {});
    const client = await connect();
    const echoed = await client.callTool({ name: "echo_args", arguments: stray });
    expect(echoed.isError).toBeFalsy();
    expect(JSON.parse(text(echoed))).toEqual({ timeZone: "UTC" });
    const now = await client.callTool({ name: "calendar_get_current_time", arguments: stray });
    expect(now.isError).toBeFalsy();
    expect(JSON.parse(text(now)).timeZone).toBe("UTC");
    const nested = await client.callTool({
      name: "nested_args",
      arguments: { range: { start: "a", end: "b", ennd: "c" } },
    });
    expect(nested.isError).toBeFalsy();
    expect(JSON.parse(text(nested))).toEqual({ range: { start: "a", end: "b" } });
    const lines = errSpy.mock.calls.filter((c) => String(c[0]).includes("GOOGLE_MCP_STRICT_PARAMS"));
    expect(lines).toHaveLength(1);
  });
});
