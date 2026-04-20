import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { google } from "googleapis";
import { z } from "zod";
import { ServiceContext } from "../../types.js";
import { textResult } from "../../utils/formatting.js";

export function registerMeetTools(server: McpServer, ctx: ServiceContext): void {
  const meetApi = () => google.meet({ version: "v2", auth: ctx.auth });
  const calApi = () => google.calendar({ version: "v3", auth: ctx.auth });

  server.tool("meet_create_link", "Create a new Google Meet meeting link", {
    summary: z.string().optional().describe("Meeting title"),
  }, async ({ summary }) => {
    const cal = calApi();
    const now = new Date();
    const later = new Date(now.getTime() + 60 * 60 * 1000);

    const res = await cal.events.insert({
      calendarId: "primary",
      conferenceDataVersion: 1,
      requestBody: {
        summary: summary || "Quick Meeting",
        start: { dateTime: now.toISOString() },
        end: { dateTime: later.toISOString() },
        conferenceData: {
          createRequest: {
            requestId: `meet-${Date.now()}`,
            conferenceSolutionKey: { type: "hangoutsMeet" },
          },
        },
      },
    });

    const meetLink = res.data.conferenceData?.entryPoints?.find((ep) => ep.entryPointType === "video")?.uri;
    return textResult({
      meetLink,
      eventId: res.data.id,
      htmlLink: res.data.htmlLink,
    });
  });

  server.tool("meet_list_meetings", "List recent Google Meet conference records", {
    pageSize: z.number().optional().describe("Number of records to return (max 100)"),
  }, async ({ pageSize }) => {
    try {
      const res = await meetApi().conferenceRecords.list({ pageSize: pageSize || 25 });
      return textResult(res.data.conferenceRecords || []);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      return textResult({ error: `Meet API error: ${msg}. Note: Meet REST API requires Google Workspace.` });
    }
  });

  server.tool("meet_get_transcript", "Get the transcript of a Google Meet recording", {
    conferenceRecordName: z.string().describe("Conference record resource name (e.g., 'conferenceRecords/abc123')"),
  }, async ({ conferenceRecordName }) => {
    try {
      const meet = meetApi();
      const transcripts = await meet.conferenceRecords.transcripts.list({ parent: conferenceRecordName });

      if (!transcripts.data.transcripts?.length) return textResult("No transcripts found for this conference.");

      const entries = await Promise.all(
        transcripts.data.transcripts.map(async (t) => {
          const entriesRes = await meet.conferenceRecords.transcripts.entries.list({ parent: t.name! });
          return entriesRes.data.transcriptEntries || [];
        })
      );

      return textResult(entries.flat());
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      return textResult({ error: `Meet API error: ${msg}. Note: Transcripts require Google Workspace Business Standard or higher.` });
    }
  });
}
