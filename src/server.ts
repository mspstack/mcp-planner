/**
 * Builds an McpServer for one session. A session is defined by the Entra app
 * credentials it uses (server-wide credentials on stdio, or client-supplied via
 * BYOK). The full tool surface is always registered; the app registration's
 * granted Graph permissions are the access control.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createRequire } from "node:module";
import type { ServerConfig } from "./config.js";
import { ToolRegistrar } from "./tools/registrar.js";
import { GraphClient, type GraphCredentials } from "./graph/client.js";
import { registerGroupTools } from "./tools/groups.js";
import { registerPlanTools } from "./tools/plans.js";
import { registerTaskTools } from "./tools/tasks.js";
import { registerAdvancedTools } from "./tools/advanced.js";

const require = createRequire(import.meta.url);
const pkg = require("../package.json") as { name: string; version: string };

export const SERVER_NAME = pkg.name;
export const SERVER_VERSION = pkg.version;

export interface SessionIdentity {
  label: string;
  credentials: GraphCredentials;
}

const INSTRUCTIONS = `# Microsoft Planner MCP server

## Finding things
- IDs chain: planner_search_groups → planner_list_plans → planner_get_plan (buckets) → planner_list_tasks → planner_get_task.
- "What is X working on" → planner_find_user → planner_list_user_tasks.
- List tools hide completed tasks by default; pass include_completed: true to see them.
- If graph_find_endpoint / graph_get are available, use them for Graph surface the curated tools don't cover — prefer the curated tool when one exists.

## Writing
- Updates handle Planner's ETag concurrency automatically; on a 412 error just retry.
- planner_update_task_details replaces the description; checklist items are addressed by item ID (from planner_get_task).
- Deletes are permanent.`;

export function createServer(config: ServerConfig, session: SessionIdentity): McpServer {
  const client = new GraphClient(session.credentials);

  const server = new McpServer(
    { name: SERVER_NAME, version: SERVER_VERSION },
    { instructions: INSTRUCTIONS }
  );

  const reg = new ToolRegistrar(server);
  registerGroupTools(reg, client);
  registerPlanTools(reg, client);
  registerTaskTools(reg, client);
  if (config.advancedToolset) {
    registerAdvancedTools(reg, client);
  }

  return server;
}
