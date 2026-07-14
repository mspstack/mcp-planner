import { describe, expect, it } from "vitest";
import {
  ALLOWED_PREFIXES,
  GraphPathError,
  findEndpoints,
  isAllowedPath,
  normalizeGraphPath,
  registerAdvancedTools,
} from "./advanced.js";
import type { ToolRegistrar, ToolSpec } from "./registrar.js";
import type { GraphClient } from "../graph/client.js";
import { loadConfig } from "../config.js";

describe("normalizeGraphPath", () => {
  it("strips the Graph host and /v1.0 prefix", () => {
    expect(normalizeGraphPath("https://graph.microsoft.com/v1.0/planner/plans/1").path).toBe(
      "/planner/plans/1"
    );
    expect(normalizeGraphPath("/v1.0/users").path).toBe("/users");
  });
  it("keeps a bare path and adds a leading slash", () => {
    expect(normalizeGraphPath("groups").path).toBe("/groups");
    expect(normalizeGraphPath("/groups").path).toBe("/groups");
  });
  it("rejects /beta paths", () => {
    expect(() => normalizeGraphPath("/beta/planner/plans/1")).toThrow(GraphPathError);
    expect(() => normalizeGraphPath("https://graph.microsoft.com/beta/users")).toThrow(GraphPathError);
  });
  it("parses an inline query (pasted @odata.nextLink) instead of dropping it", () => {
    const { path, inlineQuery } = normalizeGraphPath(
      "https://graph.microsoft.com/v1.0/groups?$select=id&$skiptoken=abc123"
    );
    expect(path).toBe("/groups");
    expect(inlineQuery).toEqual({ $select: "id", $skiptoken: "abc123" });
  });
});

describe("isAllowedPath", () => {
  it("allows the planner/groups/users prefixes", () => {
    for (const prefix of ALLOWED_PREFIXES) {
      expect(isAllowedPath(prefix)).toBe(true);
      expect(isAllowedPath(`${prefix}/anything`)).toBe(true);
    }
  });
  it("is segment-aware and rejects everything else", () => {
    expect(isAllowedPath("/groupsFoo")).toBe(false);
    expect(isAllowedPath("/me")).toBe(false);
    expect(isAllowedPath("/sites")).toBe(false);
    expect(isAllowedPath("/usersearch")).toBe(false);
  });
});

describe("findEndpoints", () => {
  it("ranks the right endpoint for a keyword query", () => {
    expect(findEndpoints("tasks assigned to a user", 3)[0]?.path).toBe("/users/{user_id}/planner/tasks");
    const bucketPaths = findEndpoints("buckets of a plan", 3).map((e) => e.path);
    expect(bucketPaths).toContain("/planner/plans/{plan_id}/buckets");
  });
  it("surfaces the query grammar doc for syntax questions", () => {
    const paths = findEndpoints("select filter odata syntax", 5).map((e) => e.path);
    expect(paths).toContain("(query grammar)");
  });
  it("returns nothing for an empty/stopword-only query", () => {
    expect(findEndpoints("", 5)).toEqual([]);
    expect(findEndpoints("how do i get the", 5)).toEqual([]);
  });
  it("respects top_k", () => {
    expect(findEndpoints("plan", 2).length).toBeLessThanOrEqual(2);
  });
});

describe("graph_get handler", () => {
  type Handler = (args: Record<string, unknown>) => Promise<{
    content: Array<{ type: "text"; text: string }>;
    isError?: boolean;
  }>;

  function capture(client: Partial<GraphClient>): Map<string, Handler> {
    const handlers = new Map<string, Handler>();
    const reg = {
      register(spec: ToolSpec, handler: Handler) {
        handlers.set(spec.name, handler);
      },
    } as unknown as ToolRegistrar;
    registerAdvancedTools(reg, client as GraphClient);
    return handlers;
  }

  const defaults = { response_format: "markdown" };

  it("rejects paths outside the allowlist without calling the client", async () => {
    let called = false;
    const handlers = capture({
      get: async () => {
        called = true;
        return {};
      },
    });
    const result = await handlers.get("graph_get")!({ ...defaults, path: "/me" });
    expect(result.isError).toBe(true);
    expect(result.content[0]!.text).toMatch(/\/planner, \/groups, \/users/);
    expect(called).toBe(false);
  });

  it("rejects /beta with a clear error", async () => {
    const handlers = capture({ get: async () => ({}) });
    const result = await handlers.get("graph_get")!({ ...defaults, path: "/beta/users" });
    expect(result.isError).toBe(true);
    expect(result.content[0]!.text).toMatch(/pinned to v1\.0/);
  });

  it("maps structured args to OData params, overriding inline ones", async () => {
    let seen: { path?: string; query?: Record<string, string | number | undefined>; headers?: Record<string, string> } = {};
    const handlers = capture({
      get: async (path: string, query?: Record<string, string | number | undefined>, headers?: Record<string, string>) => {
        seen = { path, query, headers };
        return { value: [] };
      },
    });
    await handlers.get("graph_get")!({
      ...defaults,
      path: "/groups?$select=id&$skiptoken=abc",
      select: "id,displayName",
      search: "displayName:pod",
      top: 10,
    });
    expect(seen.path).toBe("/groups");
    expect(seen.query).toMatchObject({
      $select: "id,displayName",
      $skiptoken: "abc",
      $search: '"displayName:pod"',
      $top: 10,
    });
    expect(seen.headers).toEqual({ ConsistencyLevel: "eventual" });
  });

  it("reports collection counts and surfaces the nextLink", async () => {
    const handlers = capture({
      get: async () => ({
        "@odata.context": "ctx",
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/groups?$skiptoken=next",
        value: [{ id: "1" }, { id: "2" }],
      }),
    });
    const result = await handlers.get("graph_get")!({ ...defaults, path: "/groups" });
    expect(result.isError).toBeUndefined();
    expect(result.content[0]!.text).toContain("2 record(s)");
    expect(result.content[0]!.text).toContain("$skiptoken=next");
    expect(result.content[0]!.text).not.toContain("@odata.context");
  });
});

describe("advanced toolset config flag", () => {
  const CREDS = { MS_TENANT_ID: "t", MS_CLIENT_ID: "c", MS_CLIENT_SECRET: "s" };
  it("is off by default", () => {
    expect(loadConfig([], { ...CREDS }).advancedToolset).toBe(false);
  });
  it("enables via env or flag", () => {
    expect(loadConfig([], { ...CREDS, PLANNER_ADVANCED_TOOLSET: "true" }).advancedToolset).toBe(true);
    expect(loadConfig([], { ...CREDS, PLANNER_ADVANCED_TOOLSET: "1" }).advancedToolset).toBe(true);
    expect(loadConfig(["--advanced"], { ...CREDS }).advancedToolset).toBe(true);
  });
  it("treats an unsubstituted template as off", () => {
    expect(
      loadConfig([], { ...CREDS, PLANNER_ADVANCED_TOOLSET: "${user_config.advanced}" }).advancedToolset
    ).toBe(false);
  });
});
