import { describe, expect, it } from "vitest";
import { ConfigError, loadConfig } from "./config.js";

const FULL_ENV = {
  MS_TENANT_ID: "tenant",
  MS_CLIENT_ID: "client",
  MS_CLIENT_SECRET: "secret",
};

describe("loadConfig", () => {
  it("defaults to stdio with full env credentials", () => {
    const config = loadConfig([], FULL_ENV as NodeJS.ProcessEnv);
    expect(config.transport).toBe("stdio");
    expect(config.port).toBe(3000);
    expect(config.tenantId).toBe("tenant");
  });

  it("rejects stdio without credentials", () => {
    expect(() => loadConfig([], {} as NodeJS.ProcessEnv)).toThrow(ConfigError);
  });

  it("allows http without credentials (BYOK)", () => {
    const config = loadConfig(["--transport", "http"], {} as NodeJS.ProcessEnv);
    expect(config.transport).toBe("http");
    expect(config.tenantId).toBeUndefined();
  });

  it("rejects partial credentials", () => {
    expect(() =>
      loadConfig(["--transport", "http"], { MS_TENANT_ID: "t" } as NodeJS.ProcessEnv)
    ).toThrow(/must be set together/);
  });

  it("accepts a delegated refresh token instead of a client secret", () => {
    const config = loadConfig([], {
      MS_TENANT_ID: "tenant",
      MS_CLIENT_ID: "client",
      MS_REFRESH_TOKEN: "rt-1",
    } as NodeJS.ProcessEnv);
    expect(config.refreshToken).toBe("rt-1");
    expect(config.clientSecret).toBeUndefined();
  });

  it("rejects setting both a secret and a refresh token", () => {
    expect(() =>
      loadConfig([], { ...FULL_ENV, MS_REFRESH_TOKEN: "rt-1" } as NodeJS.ProcessEnv)
    ).toThrow(/not both/);
  });

  it("rejects a refresh token without tenant/client", () => {
    expect(() =>
      loadConfig(["--transport", "http"], { MS_REFRESH_TOKEN: "rt-1" } as NodeJS.ProcessEnv)
    ).toThrow(/must be set together/);
  });

  it("treats empty-string env values as unset (MCPB hosts)", () => {
    const config = loadConfig(["--transport", "http"], {
      MS_TENANT_ID: "",
      MS_CLIENT_ID: "",
      MS_CLIENT_SECRET: "",
      PORT: "",
    } as NodeJS.ProcessEnv);
    expect(config.tenantId).toBeUndefined();
    expect(config.port).toBe(3000);
  });

  it("rejects an invalid transport and port", () => {
    expect(() => loadConfig(["--transport", "tcp"], FULL_ENV as NodeJS.ProcessEnv)).toThrow(ConfigError);
    expect(() => loadConfig(["--port", "0"], FULL_ENV as NodeJS.ProcessEnv)).toThrow(ConfigError);
  });
});
