/**
 * HTTP transport: Express app exposing
 *
 *   POST/GET/DELETE /mcp    MCP streamable-http endpoint (BYOK auth)
 *   GET  /health            liveness probe
 *
 * Authentication model — bring-your-own-credentials, two shapes:
 *  - App-only: x-ms-tenant-id + x-ms-client-id + x-ms-client-secret
 *    (client-credentials flow; the app registration's application permissions
 *    apply, writes are attributed to the app).
 *  - Delegated: x-ms-tenant-id + x-ms-client-id + x-ms-refresh-token
 *    (refresh-token grant on a public client; the signed-in USER's permissions
 *    apply and writes are attributed to them). When both a secret and a
 *    refresh token arrive, the refresh token wins — header-overlay proxies
 *    (e.g. the MCP gateway's per-user sessions) can add but not remove headers.
 *  - When the server was started with MS_TENANT_ID/MS_CLIENT_ID and a secret
 *    or refresh token, sessions may omit the headers and use the server-wide
 *    credentials instead.
 *  - The full tool surface is exposed; there is no MCP-level role gating.
 *  - A session id never carries privilege: every request re-authenticates and
 *    must present the same credentials (SHA-256 hash) the session was created with.
 */

import { createHash, randomUUID } from "node:crypto";
import express, { type Request, type Response } from "express";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import type { ServerConfig } from "../config.js";
import type { GraphCredentials } from "../graph/client.js";
import { createServer, SERVER_NAME, SERVER_VERSION } from "../server.js";

interface SessionRecord {
  transport: StreamableHTTPServerTransport;
  label: string;
  /** SHA-256 of the credentials the session was created with. */
  keyHash: string;
}

type AuthOutcome =
  | { ok: true; label: string; credentials: GraphCredentials; keyHash: string }
  | { ok: false; status: number; code: number; message: string };

const sha256 = (value: string): string => createHash("sha256").update(value).digest("hex");

function rpcError(res: Response, status: number, code: number, message: string): void {
  res.status(status).json({ jsonrpc: "2.0", error: { code, message }, id: null });
}

const unauthorized = (message: string): AuthOutcome => ({
  ok: false,
  status: 401,
  code: -32001,
  message,
});

function headerValue(req: Request, name: string): string | undefined {
  const value = req.headers[name];
  return Array.isArray(value) ? value[0] : value;
}

/** Resolve the principal for a request. Exported for tests. */
export function resolveAuth(req: Request, config: ServerConfig): AuthOutcome {
  const tenantId = headerValue(req, "x-ms-tenant-id");
  const clientId = headerValue(req, "x-ms-client-id");
  const clientSecret = headerValue(req, "x-ms-client-secret");
  const refreshToken = headerValue(req, "x-ms-refresh-token");

  // Delegated wins over app-only when both credentials arrive: header-overlay
  // proxies (the MCP gateway's per-user sessions) can add headers but never
  // remove the shared spec's, so a placeholder secret may ride along.
  if (refreshToken) {
    if (!tenantId || !clientId) {
      return unauthorized(
        "x-ms-refresh-token needs x-ms-tenant-id and x-ms-client-id alongside it."
      );
    }
    const keyHash = sha256(`${tenantId}:${clientId}:rt:${refreshToken}`);
    return {
      ok: true,
      label: `byok-user:${keyHash.slice(0, 8)}`,
      credentials: { tenantId, clientId, refreshToken },
      keyHash,
    };
  }

  const provided = [tenantId, clientId, clientSecret].filter(Boolean).length;
  if (provided > 0 && provided < 3) {
    return unauthorized(
      "Incomplete credentials: send x-ms-tenant-id + x-ms-client-id with either x-ms-client-secret (app-only) or x-ms-refresh-token (delegated)."
    );
  }

  if (provided === 3) {
    const keyHash = sha256(`${tenantId}:${clientId}:${clientSecret}`);
    return {
      ok: true,
      label: `byok:${keyHash.slice(0, 8)}`,
      credentials: { tenantId: tenantId!, clientId: clientId!, clientSecret: clientSecret! },
      keyHash,
    };
  }

  if (config.tenantId && config.clientId && (config.clientSecret || config.refreshToken)) {
    const keyHash = sha256(
      `${config.tenantId}:${config.clientId}:${config.clientSecret ?? `rt:${config.refreshToken}`}`
    );
    return {
      ok: true,
      label: "server-env",
      credentials: {
        tenantId: config.tenantId,
        clientId: config.clientId,
        ...(config.clientSecret ? { clientSecret: config.clientSecret } : {}),
        ...(config.refreshToken ? { refreshToken: config.refreshToken } : {}),
      },
      keyHash,
    };
  }

  return unauthorized(
    "Supply Entra credentials via x-ms-tenant-id + x-ms-client-id plus x-ms-client-secret (app-only) or x-ms-refresh-token (delegated)."
  );
}

function principalMatches(session: SessionRecord, auth: Extract<AuthOutcome, { ok: true }>): boolean {
  return session.keyHash === auth.keyHash;
}

export function createApp(config: ServerConfig): express.Express {
  const app = express();
  app.use(express.json({ limit: "2mb" }));

  const sessions = new Map<string, SessionRecord>();

  app.get("/health", (_req: Request, res: Response) => {
    res.json({ status: "ok", name: SERVER_NAME, version: SERVER_VERSION });
  });

  app.post("/mcp", (req: Request, res: Response) => {
    void (async () => {
      const auth = resolveAuth(req, config);
      if (!auth.ok) return rpcError(res, auth.status, auth.code, auth.message);

      const sessionId = headerValue(req, "mcp-session-id");
      if (sessionId) {
        const session = sessions.get(sessionId);
        if (!session) return rpcError(res, 404, -32000, "Session not found");
        if (!principalMatches(session, auth)) {
          return rpcError(res, 403, -32003, "Forbidden: credentials do not match this session");
        }
        await session.transport.handleRequest(req, res, req.body);
        return;
      }

      if (!isInitializeRequest(req.body)) {
        return rpcError(res, 400, -32600, "Bad request: expected an initialize request");
      }

      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (newSessionId) => {
          sessions.set(newSessionId, {
            transport,
            label: auth.label,
            keyHash: auth.keyHash,
          });
          console.error(`[auth] session ${newSessionId} created for ${auth.label}`);
        },
      });
      transport.onclose = () => {
        if (transport.sessionId) sessions.delete(transport.sessionId);
      };

      const server = createServer(config, { label: auth.label, credentials: auth.credentials });
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    })().catch((err) => {
      console.error(`[http] POST /mcp failed: ${String(err)}`);
      if (!res.headersSent) rpcError(res, 500, -32603, "Internal server error");
    });
  });

  const handleSessionRequest = (req: Request, res: Response): void => {
    void (async () => {
      const auth = resolveAuth(req, config);
      if (!auth.ok) return rpcError(res, auth.status, auth.code, auth.message);

      const sessionId = headerValue(req, "mcp-session-id");
      const session = sessionId ? sessions.get(sessionId) : undefined;
      if (!session) return rpcError(res, 404, -32000, "Session not found");
      if (!principalMatches(session, auth)) {
        return rpcError(res, 403, -32003, "Forbidden: credentials do not match this session");
      }
      await session.transport.handleRequest(req, res);
    })().catch((err) => {
      console.error(`[http] ${req.method} /mcp failed: ${String(err)}`);
      if (!res.headersSent) rpcError(res, 500, -32603, "Internal server error");
    });
  };

  app.get("/mcp", handleSessionRequest);
  app.delete("/mcp", handleSessionRequest);

  return app;
}
