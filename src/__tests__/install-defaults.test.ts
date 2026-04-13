/**
 * Tests that the install wizard pre-fills defaults from existing settings,
 * so re-running install is a quick confirmation flow.
 */
import { describe, it, expect, vi } from "vitest";

// Mock the validators module so install() doesn't make real HTTP calls
vi.mock("../validators.js", () => ({
  validateBotCredentials: vi.fn().mockResolvedValue({ ok: true }),
  validateTenant: vi.fn().mockResolvedValue({ ok: true, tenantName: "Test Tenant" }),
  discoverTeamsAndChannels: vi.fn().mockResolvedValue({ ok: false, error: "skipped" }),
  parseTeamsLink: vi.fn().mockReturnValue({}),
}));

// Mock app-package generation
vi.mock("../app-package.js", () => ({
  generateTeamsAppPackage: vi.fn().mockResolvedValue("/tmp/openacp-bot.zip"),
}));

/** Minimal InstallContext that auto-answers prompts from queued responses. */
function createTestCtx(responses: Record<string, unknown[]>) {
  const settingsData = new Map<string, unknown>();
  const terminalCalls: { method: string; args: unknown }[] = [];
  const queues = new Map<string, unknown[]>();
  for (const [m, r] of Object.entries(responses)) queues.set(m, [...r]);

  function next(method: string, args: unknown) {
    terminalCalls.push({ method, args });
    const q = queues.get(method);
    if (q && q.length) return q.shift();
    if (method === "text") return "";
    if (method === "password") return "";
    if (method === "confirm") return false;
    if (method === "select") return undefined;
    return undefined;
  }

  const noop = () => {};
  const terminal = {
    text: async (o: unknown) => next("text", o) as string,
    select: async (o: unknown) => next("select", o),
    confirm: async (o: unknown) => next("confirm", o) as boolean,
    password: async (o: unknown) => next("password", o) as string,
    multiselect: async (o: unknown) => next("multiselect", o) as unknown[],
    log: { info: noop, success: noop, warning: noop, error: noop, step: noop },
    spinner: () => ({ start: noop, stop: noop, fail: noop }),
    note: noop,
    cancel: noop,
  };

  const settings = {
    get: async (k: string) => settingsData.get(k),
    set: async (k: string, v: unknown) => { settingsData.set(k, v); },
    getAll: async () => Object.fromEntries(settingsData),
    setAll: async (all: Record<string, unknown>) => {
      settingsData.clear();
      for (const [k, v] of Object.entries(all)) settingsData.set(k, v);
    },
    delete: async (k: string) => { settingsData.delete(k); },
    clear: async () => { settingsData.clear(); },
    has: async (k: string) => settingsData.has(k),
  };

  const silentLog: any = { trace: noop, debug: noop, info: noop, warn: noop, error: noop, fatal: noop, child: () => silentLog };

  return {
    pluginName: "@hahnfeld/teams-adapter",
    terminal,
    settings,
    dataDir: "/tmp/openacp-test-data",
    log: silentLog,
    terminalCalls,
    settingsData,
  };
}

describe("install wizard defaults", () => {
  it("preserves password on empty input when existing settings exist", { timeout: 15000 }, async () => {
    const ctx = createTestCtx({
      text: ["", "", "", "", ""],       // App ID, Tenant ID, Team ID, Channel ID, notification
      password: [""],                   // Empty → should preserve existing
      select: ["single", "manual", "devtunnel"],
      confirm: [false, false, true],    // notifications=no, graph=no, defaultPort=yes
    });

    // Pre-seed existing settings
    await ctx.settings.set("botAppId", "11111111-1111-1111-1111-111111111111");
    await ctx.settings.set("botAppPassword", "existing-secret-password");
    await ctx.settings.set("tenantId", "22222222-2222-2222-2222-222222222222");
    await ctx.settings.set("teamId", "33333333-3333-3333-3333-333333333333");
    await ctx.settings.set("channelId", "19:test@thread.tacv2");
    await ctx.settings.set("botPort", 3978);
    await ctx.settings.set("tunnelMethod", "devtunnel");

    const { default: createPlugin } = await import("../plugin.js");
    const plugin = typeof createPlugin === "function" ? createPlugin() : createPlugin;
    await plugin.install!(ctx as any);

    const saved = await ctx.settings.getAll();
    expect(saved.botAppPassword).toBe("existing-secret-password");
    expect(saved.enabled).toBe(true);
  });

  it("skips 'credentials ready' confirm when settings already exist", { timeout: 15000 }, async () => {
    const ctx = createTestCtx({
      text: ["", "", "", "", ""],
      password: [""],
      select: ["single", "manual", "devtunnel"],
      confirm: [false, false, true],
    });

    await ctx.settings.set("botAppId", "11111111-1111-1111-1111-111111111111");
    await ctx.settings.set("botAppPassword", "secret");
    await ctx.settings.set("tenantId", "22222222-2222-2222-2222-222222222222");
    await ctx.settings.set("teamId", "33333333-3333-3333-3333-333333333333");
    await ctx.settings.set("channelId", "19:test@thread.tacv2");
    await ctx.settings.set("botPort", 3978);
    await ctx.settings.set("tunnelMethod", "devtunnel");

    const { default: createPlugin } = await import("../plugin.js");
    const plugin = typeof createPlugin === "function" ? createPlugin() : createPlugin;
    await plugin.install!(ctx as any);

    // The first confirm should NOT be "Do you have your Bot App ID..."
    const confirmCalls = ctx.terminalCalls.filter((c) => c.method === "confirm");
    const firstConfirm = confirmCalls[0]?.args as { message: string } | undefined;
    expect(firstConfirm?.message).not.toContain("Bot App ID and Password ready");
  });

  it("requires password on fresh install", { timeout: 15000 }, async () => {
    const ctx = createTestCtx({
      text: [
        "11111111-1111-1111-1111-111111111111",  // Bot App ID
        "22222222-2222-2222-2222-222222222222",  // Tenant ID
        "33333333-3333-3333-3333-333333333333",  // Team ID
        "19:test@thread.tacv2",                   // Channel ID
        "",                                       // Notification channel
      ],
      password: ["fresh-new-password"],
      select: ["single", "manual", "devtunnel"],
      confirm: [true, false, false, true],  // ready=yes, notifications=no, graph=no, defaultPort=yes
    });

    const { default: createPlugin } = await import("../plugin.js");
    const plugin = typeof createPlugin === "function" ? createPlugin() : createPlugin;
    await plugin.install!(ctx as any);

    const saved = await ctx.settings.getAll();
    expect(saved.botAppPassword).toBe("fresh-new-password");

    // First confirm should be the "credentials ready?" question
    const confirmCalls = ctx.terminalCalls.filter((c) => c.method === "confirm");
    const firstConfirm = confirmCalls[0]?.args as { message: string };
    expect(firstConfirm.message).toContain("Bot App ID and Password ready");
  });
});
