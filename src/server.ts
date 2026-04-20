import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ServiceContext } from "./types.js";
import { registerGmailTools } from "./services/gmail/index.js";
import { registerCalendarTools } from "./services/calendar/index.js";
import { registerMeetTools } from "./services/meet/index.js";
import { registerDriveTools } from "./services/drive/index.js";
import { registerDocsTools } from "./services/docs/index.js";
import { registerSheetsTools } from "./services/sheets/index.js";
import { registerSlidesTools } from "./services/slides/index.js";

export function createServer(ctx: ServiceContext): McpServer {
  const server = new McpServer({
    name: "google-mcp",
    version: "0.1.0",
  });

  registerGmailTools(server, ctx);
  registerCalendarTools(server, ctx);
  registerMeetTools(server, ctx);
  registerDriveTools(server, ctx);
  registerDocsTools(server, ctx);
  registerSheetsTools(server, ctx);
  registerSlidesTools(server, ctx);

  return server;
}
