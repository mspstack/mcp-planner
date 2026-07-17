import { afterEach, describe, expect, it, vi } from "vitest";
import type { Request } from "express";
import { loadConfig } from "../config.js";
import { GraphClient } from "../graph/client.js";
import { resolveAuth } from "./app.js";

const makeRequest = (headers: Record<string, string>): Request =>
  ({ headers }) as unknown as Request;

const httpConfig = loadConfig(["--transport", "http"], {} as NodeJS.ProcessEnv);

describe("resolveAuth", () => {
  it("accepts the app-only header triple", () => {
    const auth = resolveAuth(
      makeRequest({ "x-ms-tenant-id": "t", "x-ms-client-id": "c", "x-ms-client-secret": "s" }),
      httpConfig
    );
    expect(auth).toMatchObject({ ok: true, credentials: { tenantId: "t", clientId: "c", clientSecret: "s" } });
  });

  it("accepts the delegated shape (refresh token, no secret)", () => {
    const auth = resolveAuth(
      makeRequest({ "x-ms-tenant-id": "t", "x-ms-client-id": "c", "x-ms-refresh-token": "rt" }),
      httpConfig
    );
    expect(auth).toMatchObject({ ok: true, credentials: { tenantId: "t", clientId: "c", refreshToken: "rt" } });
    if (auth.ok) expect(auth.label).toMatch(/^byok-user:/);
  });

  it("prefers the refresh token when a (placeholder) secret rides along", () => {
    // Header-overlay proxies (the gateway's per-user sessions) add headers but
    // can't remove the shared spec's placeholder secret.
    const auth = resolveAuth(
      makeRequest({
        "x-ms-tenant-id": "t",
        "x-ms-client-id": "c",
        "x-ms-client-secret": "placeholder",
        "x-ms-refresh-token": "rt",
      }),
      httpConfig
    );
    expect(auth.ok).toBe(true);
    if (auth.ok) {
      expect(auth.credentials.refreshToken).toBe("rt");
      expect(auth.credentials.clientSecret).toBeUndefined();
    }
  });

  it("rejects incomplete shapes", () => {
    expect(resolveAuth(makeRequest({ "x-ms-tenant-id": "t" }), httpConfig).ok).toBe(false);
    expect(resolveAuth(makeRequest({ "x-ms-refresh-token": "rt" }), httpConfig).ok).toBe(false);
    expect(resolveAuth(makeRequest({}), httpConfig).ok).toBe(false);
  });

  it("distinct refresh tokens get distinct session key hashes", () => {
    const a = resolveAuth(
      makeRequest({ "x-ms-tenant-id": "t", "x-ms-client-id": "c", "x-ms-refresh-token": "rt-a" }),
      httpConfig
    );
    const b = resolveAuth(
      makeRequest({ "x-ms-tenant-id": "t", "x-ms-client-id": "c", "x-ms-refresh-token": "rt-b" }),
      httpConfig
    );
    if (a.ok && b.ok) expect(a.keyHash).not.toBe(b.keyHash);
  });
});

describe("GraphClient delegated token flow", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("uses the refresh grant, caches the access token, and adopts rotated refresh tokens", async () => {
    const tokenBodies: URLSearchParams[] = [];
    let issued = 0;
    vi.stubGlobal("fetch", vi.fn(async (input: string | URL, init?: RequestInit) => {
      const url = String(input);
      if (url.includes("/oauth2/v2.0/token")) {
        tokenBodies.push(new URLSearchParams(String(init?.body)));
        issued += 1;
        return new Response(
          JSON.stringify({ access_token: `at-${issued}`, expires_in: 3600, refresh_token: `rt-gen${issued + 1}` }),
          { status: 200 }
        );
      }
      return new Response(JSON.stringify({ value: [] }), { status: 200 });
    }));

    const client = new GraphClient({ tenantId: "t", clientId: "c", refreshToken: "rt-gen1" });
    await client.getList("/groups");
    await client.getList("/groups"); // cached access token — no second token call
    expect(tokenBodies.length).toBe(1);
    expect(tokenBodies[0]!.get("grant_type")).toBe("refresh_token");
    expect(tokenBodies[0]!.get("refresh_token")).toBe("rt-gen1");
    expect(tokenBodies[0]!.get("client_secret")).toBeNull();

    // Force re-auth: the rotated token (rt-gen2) must be used, not the original.
    (client as unknown as { tokenExpiresAt: number }).tokenExpiresAt = 0;
    await client.getList("/groups");
    expect(tokenBodies.length).toBe(2);
    expect(tokenBodies[1]!.get("refresh_token")).toBe("rt-gen2");
  });

  it("app-only credentials keep using client_credentials", async () => {
    const tokenBodies: URLSearchParams[] = [];
    vi.stubGlobal("fetch", vi.fn(async (input: string | URL, init?: RequestInit) => {
      const url = String(input);
      if (url.includes("/oauth2/v2.0/token")) {
        tokenBodies.push(new URLSearchParams(String(init?.body)));
        return new Response(JSON.stringify({ access_token: "at", expires_in: 3600 }), { status: 200 });
      }
      return new Response(JSON.stringify({ value: [] }), { status: 200 });
    }));

    const client = new GraphClient({ tenantId: "t", clientId: "c", clientSecret: "s" });
    await client.getList("/groups");
    expect(tokenBodies[0]!.get("grant_type")).toBe("client_credentials");
    expect(tokenBodies[0]!.get("client_secret")).toBe("s");
  });
});
