/**
 * Advanced (opt-in) toolset — an escape hatch for the Graph surface the
 * curated planner_* tools don't wrap.
 *
 *  - graph_find_endpoint: lexical search over a curated endpoint catalog.
 *  - graph_get: read-only GET passthrough, restricted to the path prefixes
 *    this server is about (/planner, /groups, /users) — defense-in-depth on
 *    top of the app registration's permissions, so a shared/BYOK app with
 *    extra grants (e.g. Mail.Read) can't be reached through this server.
 *
 * Not registered by default — enable with PLANNER_ADVANCED_TOOLSET=true or
 * --advanced. Read-only by construction (GraphClient.get is verb-locked).
 */

import { z } from "zod";
import type { ToolRegistrar } from "./registrar.js";
import type { GraphClient } from "../graph/client.js";
import { GRAPH_ENDPOINTS, type EndpointDoc } from "../reference/graph-endpoints.js";
import { clip, failure, json, responseFormatField, text } from "./shared.js";

/** Path prefixes graph_get may reach. Exported for tests; extend deliberately. */
export const ALLOWED_PREFIXES = ["/planner", "/groups", "/users"] as const;

export function isAllowedPath(path: string): boolean {
  return ALLOWED_PREFIXES.some((prefix) => path === prefix || path.startsWith(`${prefix}/`));
}

/** Thrown by normalizeGraphPath for paths this server refuses outright. */
export class GraphPathError extends Error {}

export interface NormalizedGraphPath {
  path: string;
  /** Query params parsed from an inline `?…` (e.g. a pasted @odata.nextLink). */
  inlineQuery: Record<string, string>;
}

/**
 * Normalize a user/model-supplied path. Unlike the ConnectWise sibling this
 * PARSES an inline query instead of dropping it, so a pasted @odata.nextLink
 * (with its $skiptoken) works directly as `path`; structured args win over
 * inline ones. /beta is rejected — the client is pinned to v1.0.
 */
export function normalizeGraphPath(raw: string): NormalizedGraphPath {
  let s = raw.trim().replace(/^https?:\/\/graph\.microsoft\.com/i, "");
  const inlineQuery: Record<string, string> = {};
  const qi = s.indexOf("?");
  if (qi >= 0) {
    for (const [key, value] of new URLSearchParams(s.slice(qi + 1))) inlineQuery[key] = value;
    s = s.slice(0, qi);
  }
  if (!s.startsWith("/")) s = `/${s}`;
  if (/^\/beta(\/|$)/i.test(s)) {
    throw new GraphPathError("Graph /beta endpoints are not supported — this server is pinned to v1.0.");
  }
  s = s.replace(/^\/v1\.0(?=\/|$)/i, "");
  if (!s.startsWith("/")) s = `/${s}`;
  return { path: s, inlineQuery };
}

// Lexical endpoint search — kept in sync with mcp-connectwise-psa src/tools/advanced.ts.

const STOP = new Set(["the", "a", "an", "of", "for", "to", "and", "or", "in", "on", "by", "with", "how", "do", "i", "get", "list", "all", "graph", "planner"]);

function tokenize(s: string): string[] {
  return s
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((t) => t.length > 1 && !STOP.has(t));
}

/** Lexical relevance score of an endpoint for a set of query tokens. */
function scoreEndpoint(doc: EndpointDoc, tokens: string[]): number {
  const strong = `${doc.path} ${doc.summary}`.toLowerCase();
  const medium = `${doc.module} ${doc.keyParams ?? ""} ${doc.commonFields ?? ""}`.toLowerCase();
  const weak = `${doc.notes ?? ""} ${doc.coveredBy ?? ""}`.toLowerCase();
  let score = 0;
  for (const t of tokens) {
    if (strong.includes(t)) score += 3;
    else if (medium.includes(t)) score += 2;
    else if (weak.includes(t)) score += 1;
  }
  return score;
}

/** Rank the catalog for a query. Pure — exported for tests. */
export function findEndpoints(query: string, topK = 5): EndpointDoc[] {
  const tokens = tokenize(query);
  if (tokens.length === 0) return [];
  return GRAPH_ENDPOINTS.map((doc) => ({ doc, score: scoreEndpoint(doc, tokens) }))
    .filter((r) => r.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, topK)
    .map((r) => r.doc);
}

function endpointBlock(d: EndpointDoc): string {
  const lines = [`### ${d.path}  \`${d.methods}\`  _(${d.module})_`, d.summary];
  if (d.keyParams) lines.push(`- **Params**: ${d.keyParams}`);
  if (d.commonFields) lines.push(`- **Fields**: ${d.commonFields}`);
  if (d.coveredBy) lines.push(`- **Prefer the curated tool(s)**: ${d.coveredBy}`);
  if (d.notes) lines.push(`- _${d.notes}_`);
  return lines.join("\n");
}

export function registerAdvancedTools(reg: ToolRegistrar, client: GraphClient): void {
  reg.register(
    {
      name: "graph_find_endpoint",
      title: "Find a Microsoft Graph Endpoint",
      description:
        "Discover which Microsoft Graph endpoint (and its query params) fits a task — search a curated catalog " +
        "of the /planner, /groups and /users surface. Use the result with graph_get, or prefer the named " +
        "curated tool when one exists.",
      inputSchema: {
        query: z.string().min(1).describe('What you want to do (e.g. "task board order", "group owners")'),
        top_k: z.number().int().positive().max(20).default(5).describe("How many endpoints to return"),
        response_format: responseFormatField,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: { query: string; top_k: number; response_format: "markdown" | "json" }) => {
      const hits = findEndpoints(args.query, args.top_k);
      if (hits.length === 0) return text(`No endpoint matched "${args.query}". Try broader keywords (resource name).`);
      if (args.response_format === "json") return text(clip(json(hits)));
      const lines = [`# Endpoints for "${args.query}"`, "", ...hits.map(endpointBlock), "", "Call one with **graph_get** (read-only)."];
      return text(clip(lines.join("\n\n")));
    }
  );

  reg.register(
    {
      name: "graph_get",
      title: "Call any Graph GET Endpoint (planner/groups/users)",
      description:
        "Read-only GET for any Microsoft Graph v1.0 path under /planner, /groups or /users. ALWAYS pass select " +
        "to keep responses small. NOTE: planner collections ignore filter/top server-side. A full @odata.nextLink " +
        "URL is accepted as path. Discover paths with graph_find_endpoint.",
      inputSchema: {
        path: z.string().min(1).describe('Graph v1.0 path, e.g. "/planner/plans/{id}/tasks" — or a full @odata.nextLink URL'),
        select: z.string().optional().describe("$select — comma-separated properties (strongly recommended)"),
        filter: z.string().optional().describe("$filter — directory objects only; NOT supported on planner collections"),
        expand: z.string().optional().describe("$expand — related resources to inline"),
        orderby: z.string().optional().describe("$orderby — directory objects only"),
        top: z.number().int().positive().max(999).optional().describe("$top — NOT supported on planner collections"),
        skiptoken: z.string().optional().describe("$skiptoken from a previous page (or just pass the whole nextLink as path)"),
        search: z.string().optional().describe('$search, e.g. "displayName:pod" — directory objects only; the ConsistencyLevel header is added automatically'),
        response_format: responseFormatField,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    },
    async (args: {
      path: string;
      select?: string;
      filter?: string;
      expand?: string;
      orderby?: string;
      top?: number;
      skiptoken?: string;
      search?: string;
      response_format: "markdown" | "json";
    }) => {
      try {
        const { path, inlineQuery } = normalizeGraphPath(args.path);
        if (!isAllowedPath(path)) {
          return failure(
            new Error(`graph_get only reaches paths under ${ALLOWED_PREFIXES.join(", ")} — got "${path}".`)
          );
        }
        const query: Record<string, string | number | undefined> = { ...inlineQuery };
        if (args.select) query.$select = args.select;
        if (args.filter) query.$filter = args.filter;
        if (args.expand) query.$expand = args.expand;
        if (args.orderby) query.$orderby = args.orderby;
        if (args.top !== undefined) query.$top = args.top;
        if (args.skiptoken) query.$skiptoken = args.skiptoken;
        if (args.search) query.$search = args.search.startsWith('"') ? args.search : `"${args.search}"`;
        const result = await client.get<Record<string, unknown>>(
          path,
          query,
          args.search ? { ConsistencyLevel: "eventual" } : undefined
        );
        const { "@odata.context": _context, "@odata.nextLink": nextLink, ...rest } = result ?? {};
        const value = rest.value;
        const isCollection = Array.isArray(value);
        const body = isCollection ? { items: value } : rest;
        const count = isCollection ? `${value.length} record(s)` : "1 record";
        const footer =
          typeof nextLink === "string"
            ? `\n\n---\nMore available — pass this nextLink as path:\n${nextLink}`
            : "";
        return text(clip(`GET ${path} — ${count}\n\n${json(body)}${footer}`, "Narrow with select, or page via the nextLink."));
      } catch (error) {
        return failure(error);
      }
    }
  );
}
