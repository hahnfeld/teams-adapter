import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";
import type { CommandRegistry } from "@openacp/plugin-sdk";

/**
 * Handle /bypass — toggle auto-approve permissions for the current session.
 * Mirrors Telegram's dangerous mode toggle pattern.
 */
export async function handleBypass(ctx: CommandContext): Promise<void> {
  if (!ctx.sessionId) {
    await ctx.reply("❌ No active session. Bypass applies per-session.");
    return;
  }

  const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
  if (!session) {
    await ctx.reply("❌ Session not found.");
    return;
  }

  // Toggle via command registry if available (propagates through ACP)
  const registry = ctx.adapter.core.lifecycleManager?.serviceRegistry?.get<CommandRegistry>("command-registry");
  if (registry) {
    const currentBypass = !!session.clientOverrides?.bypassPermissions;
    const newState = !currentBypass;
    try {
      await registry.execute(newState ? "/bypass_permissions on" : "/bypass_permissions off", {
        raw: newState ? "on" : "off",
        sessionId: ctx.sessionId,
        channelId: "teams",
        userId: ctx.userId,
        reply: async () => {},
      });
      const icon = newState ? "☠️" : "🔐";
      const label = newState ? "Bypass enabled — permissions auto-approved" : "Bypass disabled — approvals required";
      await ctx.reply(`${icon} ${label}`);
      return;
    } catch { /* fall through */ }
  }

  // Direct fallback: toggle client override
  const currentBypass = !!session.clientOverrides?.bypassPermissions;
  session.clientOverrides = { ...session.clientOverrides, bypassPermissions: !currentBypass };
  const icon = !currentBypass ? "☠️" : "🔐";
  const label = !currentBypass ? "Bypass enabled — permissions auto-approved" : "Bypass disabled — approvals required";
  await ctx.reply(`${icon} ${label}`);
}

/**
 * Handle /tts [on|off] — toggle text-to-speech for the current session.
 */
export async function handleTTS(ctx: CommandContext, mode?: string): Promise<void> {
  if (!ctx.sessionId) {
    await ctx.reply("❌ No active session.");
    return;
  }

  const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
  if (!session) {
    await ctx.reply("❌ Session not found.");
    return;
  }

  if (mode === "on" || mode === "off") {
    session.voiceMode = mode;
    await ctx.reply(`🔊 TTS ${mode === "on" ? "enabled" : "disabled"}`);
    return;
  }

  // Toggle
  const newMode = session.voiceMode === "on" ? "off" : "on";
  session.voiceMode = newMode;
  await ctx.reply(`🔊 TTS ${newMode === "on" ? "enabled" : "disabled"}`);
}

/**
 * Handle /restart — restart the OpenACP assistant session.
 */
export async function handleRestart(ctx: CommandContext): Promise<void> {
  try {
    await ctx.adapter.restartAssistant();
    await ctx.reply("🔄 OpenACP assistant restarting...");
  } catch (err) {
    log.error({ err }, "[admin] restartAssistant failed");
    await ctx.reply("❌ Restart failed. Check logs for details.");
  }
}

/**
 * Handle /respawn — restart the assistant session.
 */
export async function handleRespawn(ctx: CommandContext): Promise<void> {
  try {
    await ctx.adapter.respawnAssistant();
    await ctx.reply("🔄 Assistant session restarted.");
  } catch (err) {
    log.error({ err }, "[admin] respawnAssistant failed");
    await ctx.reply("❌ Respawn failed. Check logs for details.");
  }
}

/**
 * Handle /update — check for and apply updates.
 */
export async function handleUpdate(ctx: CommandContext): Promise<void> {
  const registry = ctx.adapter.core.lifecycleManager?.serviceRegistry?.get<CommandRegistry>("command-registry");
  if (registry) {
    try {
      const response = await registry.execute("/update", {
        raw: "",
        sessionId: null,
        channelId: "teams",
        userId: ctx.userId,
        reply: async (content: string) => { await ctx.reply(content); },
      });
      if (response.type === "text") {
        await ctx.reply(response.text);
      }
      return;
    } catch { /* fall through */ }
  }
  await ctx.reply("📦 Update check not available. Run `openacp update` from the terminal.");
}

/**
 * Handle /outputmode <level> [scope] — set output detail level.
 */
export async function handleOutputMode(
  ctx: CommandContext,
  level?: string,
  scope?: string,
): Promise<void> {
  if (!level || level === "reset") {
    if (ctx.sessionId) {
      ctx.adapter.setSessionOutputMode(ctx.sessionId, "medium");
    }
    await ctx.reply("🔄 Output mode reset to **medium**");
    return;
  }

  if (level !== "low" && level !== "medium" && level !== "high") {
    await ctx.reply("⚠️ Valid levels: `low`, `medium`, `high`, `reset`");
    return;
  }

  if (scope === "session" && ctx.sessionId) {
    try {
      await ctx.adapter.core.sessionManager.patchRecord(ctx.sessionId, { outputMode: level } as any);
    } catch { /* best effort */ }
    ctx.adapter.setSessionOutputMode(ctx.sessionId, level as "low" | "medium" | "high");
    await ctx.reply(`🔄 Output mode set to **${level}** for this session`);
  } else if (ctx.sessionId) {
    ctx.adapter.setSessionOutputMode(ctx.sessionId, level as "low" | "medium" | "high");
    await ctx.reply(`🔄 Output mode set to **${level}** for this session`);
  } else {
    await ctx.reply(`🔄 Output mode set to **${level}** (use in a session for per-session control)`);
  }
}
