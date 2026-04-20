import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { OAuth2Client } from "google-auth-library";

export interface ServiceContext {
  auth: OAuth2Client;
}

export type RegisterTools = (server: McpServer, ctx: ServiceContext) => void;
