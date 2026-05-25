#!/usr/bin/env node
import { program } from "commander";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { loadAuth } from "./auth.js";
import { createServer } from "./server.js";
import { homedir } from "node:os";
import { join } from "node:path";

program
  .name("google-mcp")
  .description("Consolidated Google MCP server")
  .requiredOption("--slug <slug>", "Google account slug (e.g. jane-acme-com)")
  .option(
    "--token-dir <dir>",
    "Directory containing credentials files",
    join(homedir(), ".config", "openbrain", "tokens")
  )
  .parse();

const opts = program.opts<{ slug: string; tokenDir: string }>();

// The parent (Claude Code) terminates stdio MCP servers with SIGINT/SIGTERM
// during session teardown. Without these handlers gVisor (Cloud Run's sandbox)
// reports each as `Uncaught signal: 2` at ERROR severity in Cloud Logging.
const shutdown = (): never => process.exit(0);
process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);

const auth = loadAuth(opts.slug, opts.tokenDir);
const server = createServer({ auth });
const transport = new StdioServerTransport();
await server.connect(transport);
