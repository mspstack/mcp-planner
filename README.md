# mcp-planner

MCP server for **Microsoft Planner** via the Microsoft Graph API. Find groups and plans, list buckets and tasks, and create, update, assign, complete, or delete tasks — including descriptions and checklists — with Planner's ETag concurrency handled automatically.

Sibling project to [mcp-itglue](https://github.com/selic/mcp-itglue) and [mcp-connectwise-psa](https://github.com/selic/mcp-connectwise-psa) — same architecture.

## Tools

| Tool | Description |
|---|---|
| `planner_search_groups` | Find Microsoft 365 groups (Teams) by name → group ID |
| `planner_find_user` | Find a user by name/UPN → user ID for assignments |
| `planner_list_plans` | List plans owned by a group |
| `planner_get_plan` | Plan + its buckets |
| `planner_create_bucket` | Create a bucket in a plan |
| `planner_list_tasks` | Tasks in a plan or bucket (filter by assignee, open/completed) |
| `planner_list_user_tasks` | All tasks assigned to a user, across plans |
| `planner_get_task` | Task with description and checklist |
| `planner_create_task` | Create task (bucket, due date, priority, assignees, description) |
| `planner_update_task` | Update title/bucket/due/priority/progress/assignees |
| `planner_update_task_details` | Update description; add or (un)check checklist items |
| `planner_delete_task` | Permanently delete a task |
| `graph_find_endpoint` † | Search a curated catalog of the /planner, /groups, /users Graph surface |
| `graph_get` † | Read-only GET for any Graph v1.0 path under /planner, /groups, /users |

† Advanced toolset (opt-in, off by default) — an escape hatch for Graph surface the curated tools don't wrap. Enable with `PLANNER_ADVANCED_TOOLSET=true` or `--advanced`. `graph_get` is verb-locked to GET, rejects `/beta`, and only reaches the three path prefixes above, so a shared app registration's other permissions (e.g. mail) stay out of reach.

## Setup

### 1. Entra ID app registration

1. [Entra admin center](https://entra.microsoft.com) → App registrations → **New registration**
2. API permissions → **Application permissions** → add `Tasks.ReadWrite.All`, `GroupMember.Read.All`, `User.Read.All` → **Grant admin consent**
3. Certificates & secrets → **New client secret** — note the value

### 2. Run

```bash
MS_TENANT_ID=<tenant> MS_CLIENT_ID=<client-id> MS_CLIENT_SECRET=<secret> npx -y mcp-planner
```

Claude Code:

```bash
claude mcp add planner --env MS_TENANT_ID=<tenant> --env MS_CLIENT_ID=<client-id> --env MS_CLIENT_SECRET=<secret> -- npx -y mcp-planner
```

### HTTP mode

```bash
npx -y mcp-planner --transport http --port 3000
```

Sessions authenticate per-request with `x-ms-tenant-id` / `x-ms-client-id` / `x-ms-client-secret` headers (BYOK), or fall back to the `MS_*` environment credentials when set. Health probe at `GET /health`.

### Docker

```bash
docker build -t mcp-planner .
docker run -p 3000:3000 -e MS_TENANT_ID=... -e MS_CLIENT_ID=... -e MS_CLIENT_SECRET=... mcp-planner
```

## Access model

No MCP-level role gating: the Entra app registration's granted Graph permissions **are** the access control. Point sessions at different app registrations (BYOK headers) to scope what they can do.

## Notes

- Planner requires an `If-Match` ETag on every update/delete — the tools fetch the current resource and pass its ETag automatically. On a 412 (concurrent change), just retry.
- Priority mapping: urgent=1, important=3, medium=5, low=9 (Graph uses 0–10).
- Progress: not started (0), in progress (50), completed (100).

## Development

```bash
npm install
npm run dev        # stdio
npm run dev:http   # http
npm test
npm run build
npm run bundle     # Claude Desktop .mcpb
```

## License

MIT
