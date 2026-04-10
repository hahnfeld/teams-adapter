import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { validateBotCredentials, validateTenant, parseTeamsLink } from "../validators.js";

// Mock global fetch
const originalFetch = globalThis.fetch;

describe("validateBotCredentials", () => {
  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it("returns ok:true when token acquisition succeeds", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve({ access_token: "test-token", expires_in: 3600 }),
    });

    const result = await validateBotCredentials("app-id", "password");
    expect(result.ok).toBe(true);
  });

  it("returns invalid_client error for wrong credentials", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 401,
      json: () => Promise.resolve({ error: "invalid_client", error_description: "bad creds" }),
    });

    const result = await validateBotCredentials("app-id", "wrong-password");
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error).toContain("Invalid App ID or Password");
    }
  });

  it("returns unauthorized_client error", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 403,
      json: () => Promise.resolve({ error: "unauthorized_client" }),
    });

    const result = await validateBotCredentials("app-id", "password");
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error).toContain("not authorized");
    }
  });

  it("handles network errors gracefully", async () => {
    globalThis.fetch = vi.fn().mockRejectedValue(new Error("ECONNREFUSED"));

    const result = await validateBotCredentials("app-id", "password");
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error).toContain("Network error");
    }
  });

  it("uses tenantId when provided", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve({ access_token: "test", expires_in: 3600 }),
    });

    await validateBotCredentials("app-id", "password", "my-tenant-id");

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock.calls[0];
    expect(fetchCall[0]).toContain("my-tenant-id");
  });

  it("uses botframework.com when no tenantId", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve({ access_token: "test", expires_in: 3600 }),
    });

    await validateBotCredentials("app-id", "password");

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock.calls[0];
    expect(fetchCall[0]).toContain("botframework.com");
  });
});

describe("validateTenant", () => {
  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it("returns ok:true when tenant is valid", async () => {
    globalThis.fetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "test", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ issuer: "https://login.microsoftonline.com/my-tenant/v2.0" }),
      });

    const result = await validateTenant("app-id", "password", "my-tenant");
    expect(result.ok).toBe(true);
  });

  it("returns ok:false when credentials fail for tenant", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 401,
      json: () => Promise.resolve({ error: "invalid_client" }),
    });

    const result = await validateTenant("app-id", "bad-pass", "my-tenant");
    expect(result.ok).toBe(false);
  });
});

describe("parseTeamsLink", () => {
  it("extracts teamId and channelId from a Teams link", () => {
    const link = "https://teams.microsoft.com/l/channel/19%3Aabc123%40thread.tacv2/General?groupId=team-guid-123&tenantId=tenant-guid-456";
    const result = parseTeamsLink(link);
    expect(result.teamId).toBe("team-guid-123");
    expect(result.channelId).toBe("19:abc123@thread.tacv2");
    expect(result.tenantId).toBe("tenant-guid-456");
  });

  it("returns empty object for invalid URLs", () => {
    const result = parseTeamsLink("not a url");
    expect(result).toEqual({});
  });

  it("handles links without tenantId", () => {
    const link = "https://teams.microsoft.com/l/channel/19%3Axyz%40thread.tacv2/Dev?groupId=abc-123";
    const result = parseTeamsLink(link);
    expect(result.teamId).toBe("abc-123");
    expect(result.channelId).toBe("19:xyz@thread.tacv2");
    expect(result.tenantId).toBeUndefined();
  });
});
