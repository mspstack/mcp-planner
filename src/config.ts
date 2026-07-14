/**
 * Server configuration, resolved from CLI flags and environment variables.
 *
 * Environment variables:
 *   MS_TENANT_ID     Entra ID (Azure AD) tenant id (required for stdio)
 *   MS_CLIENT_ID     App registration client id (required for stdio)
 *   MS_CLIENT_SECRET App registration client secret (required for stdio)
 *   TRANSPORT        stdio | http (default: stdio)
 *   PORT             HTTP port (default: 3000)
 *   PLANNER_ADVANCED_TOOLSET
 *                    true|1 registers the advanced toolset (graph_get /
 *                    graph_find_endpoint) — also via the --advanced flag.
 *                    Off by default; the tools are an escape hatch for Graph
 *                    surface the curated tools don't wrap.
 *
 * Access model: stdio uses the server-wide app credentials above (single local
 * user / single tenant). HTTP sessions each bring their own app credentials
 * (BYOK) via x-ms-tenant-id / x-ms-client-id / x-ms-client-secret headers;
 * Microsoft Graph enforces the app's granted permissions (client-credentials
 * flow, application permissions: Tasks.ReadWrite.All, GroupMember.Read.All,
 * User.Read.All). There is no MCP-level role gating.
 */

export type Transport = "stdio" | "http";

export interface ServerConfig {
  transport: Transport;
  port: number;
  /** Server-wide app credentials; used by stdio, absent on HTTP (BYOK). */
  tenantId: string | undefined;
  clientId: string | undefined;
  clientSecret: string | undefined;
  /** Register the advanced toolset (graph_get / graph_find_endpoint). */
  advancedToolset: boolean;
}

export class ConfigError extends Error {}

function flagValue(argv: string[], name: string): string | undefined {
  const idx = argv.indexOf(name);
  if (idx === -1) return undefined;
  const value = argv[idx + 1];
  if (value === undefined || value.startsWith("--")) {
    throw new ConfigError(`Missing value for ${name}`);
  }
  return value;
}

export function loadConfig(
  argv: string[] = process.argv.slice(2),
  env: NodeJS.ProcessEnv = process.env
): ServerConfig {
  // `||` not `??`: desktop hosts (MCPB) pass unset optional config as empty strings
  const transport = ((flagValue(argv, "--transport") ?? env.TRANSPORT) || "stdio") as Transport;
  if (transport !== "stdio" && transport !== "http") {
    throw new ConfigError(`Invalid transport "${transport}" — expected "stdio" or "http"`);
  }

  const portRaw = (flagValue(argv, "--port") ?? env.PORT) || "3000";
  const port = Number(portRaw);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new ConfigError(`Invalid port "${portRaw}"`);
  }

  const tenantId = env.MS_TENANT_ID || undefined;
  const clientId = env.MS_CLIENT_ID || undefined;
  const clientSecret = env.MS_CLIENT_SECRET || undefined;

  const provided = [tenantId, clientId, clientSecret].filter(Boolean).length;
  if (provided > 0 && provided < 3) {
    throw new ConfigError("MS_TENANT_ID, MS_CLIENT_ID and MS_CLIENT_SECRET must be set together");
  }

  if (transport === "stdio" && !tenantId) {
    throw new ConfigError(
      "MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET are required for stdio transport"
    );
  }

  // `||` keeps unset/empty/unsubstituted-template values falsy here too
  const advancedEnv = ((env.PLANNER_ADVANCED_TOOLSET || "").includes("${")
    ? ""
    : env.PLANNER_ADVANCED_TOOLSET || ""
  ).toLowerCase();
  const advancedToolset = argv.includes("--advanced") || advancedEnv === "true" || advancedEnv === "1";

  return { transport, port, tenantId, clientId, clientSecret, advancedToolset };
}
