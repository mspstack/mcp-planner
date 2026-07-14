# CLAUDE.md

MCP server for Microsoft Planner via Microsoft Graph. Sibling project to
mcp-connectwise-psa (copied from it) and mcp-itglue — same architecture.

## Commands

- `npm run build` — tsc to dist/
- `npm test` — vitest
- `npm run dev` / `npm run dev:http` — tsx, stdio or HTTP transport
- `npm run bundle` — build the Claude Desktop `.mcpb`

## Architecture

- `src/index.ts` — CLI entry; `--transport stdio|http`, `--port`; USAGE text
- `src/config.ts` — `loadConfig(argv, env)` → `ServerConfig`; `ConfigError`; MS_* env vars
- `src/server.ts` — `createServer(session)` builds a per-session `McpServer`; `INSTRUCTIONS` string
- `src/graph/client.ts` — `GraphClient`: client-credentials token (cached, 60s skew),
  `get/getList/post/patch/delete`; PATCH/DELETE take an explicit `etag` (If-Match);
  PATCH sends `Prefer: return=representation`; `GraphApiError` + `describeError`
- `src/graph/types.ts` — Graph resource shapes; `priorityLabel`/`PRIORITY_VALUES`/`progressLabel`
- `src/http/app.ts` — Express app; BYOK via `x-ms-tenant-id`/`x-ms-client-id`/`x-ms-client-secret`
  headers with fallback to server env creds; session principal pinned by SHA-256 hash
- `src/tools/registrar.ts` — `ToolRegistrar`; no MCP-level role gating (Graph is the access control)
- `src/tools/shared.ts` — zod fragments (`responseFormatField`, `topField`), `text/failure/json/clip`
- `src/tools/{groups,plans,tasks}.ts` — `planner_*` tools
- `src/tools/advanced.ts` — opt-in escape hatch (`PLANNER_ADVANCED_TOOLSET=true` / `--advanced`, off by default):
  `graph_get` (read-only GET via the verb-locked `GraphClient.get`, allowlisted to `/planner`, `/groups`,
  `/users`; `/beta` rejected; a pasted `@odata.nextLink` works as `path` — its inline query is parsed and
  merged, structured args win) and `graph_find_endpoint` (lexical search over the hand-maintained
  `src/reference/graph-endpoints.ts`). Search functions kept in sync with mcp-connectwise-psa's `advanced.ts`

## Graph/Planner gotchas

- Every planner PATCH/DELETE requires `If-Match` with the resource's `@odata.etag`;
  tools GET first, then write. 412 = concurrent modification, retry.
- Task details (description, checklist) are a separate resource
  (`/planner/tasks/{id}/details`) with their **own** etag.
- `percentComplete` only meaningful as 0 / 50 / 100. `priority` 0–10
  (urgent=1, important=3, medium=5, low=9).
- Assignments/checklist are open dictionaries; entries need
  `"@odata.type": "#microsoft.graph.plannerAssignment"` / `plannerChecklistItem`
  and `orderHint: " !"`; set an entry to `null` to remove it.
- Group `$search` requires the `ConsistencyLevel: eventual` header.
- App-only (client credentials) cannot use `/me` — user-scoped reads go through
  `/users/{id}/planner/tasks`.
- Planner collections don't support `$filter`/`$top` server-side — filtering is client-side.

## Conventions

- Tool responses via `shared.ts` helpers; markdown default, `response_format: "json"` toggle;
  responses clipped at 25k chars.
- All logging via `console.error` (stdout is reserved for the stdio transport).
- Never log secrets — labels/hashes only.
