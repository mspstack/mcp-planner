#!/usr/bin/env node
/** CLI entry point — runs the MCP server over stdio (default) or HTTP. */

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { ConfigError, loadConfig, type ServerConfig } from "./config.js";
import { createServer, SERVER_NAME, SERVER_VERSION } from "./server.js";
import { createApp } from "./http/app.js";

const USAGE = `${SERVER_NAME} v${SERVER_VERSION}

Usage: mcp-planner [options]

Options:
  --transport stdio|http   Transport (default: stdio; env TRANSPORT)
  --port <n>               HTTP port (default: 3000; env PORT)
  --advanced               Register the advanced toolset: graph_get /
                           graph_find_endpoint (env PLANNER_ADVANCED_TOOLSET)
  --help                   Show this help

Environment:
  MS_TENANT_ID             Entra ID tenant id (required for stdio)
  MS_CLIENT_ID             App registration client id (required for stdio)
  MS_CLIENT_SECRET         App registration client secret (required for stdio)

The app registration needs application permissions with admin consent:
Tasks.ReadWrite.All, GroupMember.Read.All, User.Read.All.

HTTP sessions may authenticate per-request with their own app credentials via
the x-ms-tenant-id / x-ms-client-id / x-ms-client-secret headers (BYOK), or
fall back to the MS_* environment credentials when those are set.
`;

function logStartupSummary(config: ServerConfig): void {
  if (config.tenantId) {
    console.error(
      "[auth] HTTP: sessions without x-ms-* headers use the server-wide MS_* credentials."
    );
  } else {
    console.error(
      "[auth] HTTP: each session must present x-ms-tenant-id / x-ms-client-id / x-ms-client-secret (BYOK)."
    );
  }
}

async function runStdio(config: ServerConfig): Promise<void> {
  const server = createServer(config, {
    label: "stdio",
    credentials: {
      tenantId: config.tenantId!,
      clientId: config.clientId!,
      clientSecret: config.clientSecret!,
    },
  });
  await server.connect(new StdioServerTransport());
  console.error(`${SERVER_NAME} v${SERVER_VERSION} running on stdio (tenant ${config.tenantId})`);
}

async function runHttp(config: ServerConfig): Promise<void> {
  logStartupSummary(config);
  const app = createApp(config);
  await new Promise<void>((resolve) => {
    app.listen(config.port, "0.0.0.0", () => resolve());
  });
  console.error(`${SERVER_NAME} v${SERVER_VERSION} listening on :${config.port}`);
}

async function main(): Promise<void> {
  const argv = process.argv.slice(2);
  if (argv.includes("--help") || argv.includes("-h")) {
    console.log(USAGE);
    return;
  }

  let config: ServerConfig;
  try {
    config = loadConfig(argv);
  } catch (err) {
    if (err instanceof ConfigError) {
      console.error(`Configuration error: ${err.message}\n`);
      console.error(USAGE);
      process.exit(1);
    }
    throw err;
  }

  if (config.transport === "http") await runHttp(config);
  else await runStdio(config);
}

main().catch((err) => {
  console.error(`Fatal: ${String(err)}`);
  process.exit(1);
});
