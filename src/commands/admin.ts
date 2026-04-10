import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";

export async function handleBypass(ctx: CommandContext): Promise<void> {
  await ctx.reply("🔓 Bypass mode toggled (not yet implemented)");
}

export async function handleTTS(ctx: CommandContext, mode?: string): Promise<void> {
  await ctx.reply(`🔊 TTS mode: ${mode ?? "toggle"} (not yet implemented)`);
}

export async function handleRestart(ctx: CommandContext): Promise<void> {
  try {
    await ctx.adapter.restartAssistant();
    await ctx.reply("🔄 OpenACP is restarting...");
  } catch (err) {
    log.error({ err }, "[admin] restartAssistant failed");
    await ctx.reply("❌ Restart failed. Check logs for details.");
  }
}

export async function handleRespawn(ctx: CommandContext): Promise<void> {
  try {
    await ctx.adapter.respawnAssistant();
    await ctx.reply("🔄 Assistant session restarted.");
  } catch (err) {
    log.error({ err }, "[admin] respawnAssistant failed");
    await ctx.reply("❌ Respawn failed. Check logs for details.");
  }
}

export async function handleUpdate(ctx: CommandContext): Promise<void> {
  await ctx.reply("📦 Update check... (not yet implemented)");
}

export async function handleOutputMode(
  ctx: CommandContext,
  level?: string,
  scope?: string,
): Promise<void> {
  if (!level || level === "reset") {
    await ctx.reply("🔄 Resetting output mode to default (not yet implemented)");
    return;
  }
  if (level !== "low" && level !== "medium" && level !== "high") {
    await ctx.reply("⚠️ Valid levels: low, medium, high, reset");
    return;
  }

  if (scope === "session" && ctx.sessionId) {
    await ctx.adapter.core.sessionManager.patchRecord(ctx.sessionId, { outputMode: level } as any);
    ctx.adapter.setSessionOutputMode(ctx.sessionId, level as "low" | "medium" | "high");
    await ctx.reply(`🔄 Output mode set to **${level}** for this session`);
  } else {
    // Adapter-level default
    await ctx.reply(`🔄 Output mode set to **${level}** (not yet persisted to config)`);
  }
}