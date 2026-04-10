import type { CommandContext } from "./index.js";

export async function handleBypass(ctx: CommandContext): Promise<void> {
  await ctx.reply("🔓 Bypass mode toggled (not yet implemented)");
}

export async function handleTTS(ctx: CommandContext, mode?: string): Promise<void> {
  await ctx.reply(`🔊 TTS mode: ${mode ?? "toggle"} (not yet implemented)`);
}

export async function handleRestart(ctx: CommandContext): Promise<void> {
  await ctx.reply("🔄 Restarting OpenACP... (not yet implemented)");
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
    await ctx.reply(`🔄 Output mode set to **${level}** for this session`);
  } else {
    // Adapter-level default
    await ctx.reply(`🔄 Output mode set to **${level}** (not yet persisted to config)`);
  }
}