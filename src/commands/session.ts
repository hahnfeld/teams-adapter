import type { CommandContext } from "./index.js";

export async function handleCancel(ctx: CommandContext): Promise<void> {
  if (!ctx.sessionId) {
    await ctx.reply("❌ No active session to cancel.");
    return;
  }
  const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
  if (!session) {
    await ctx.reply("❌ Session not found.");
    return;
  }
  await session.destroy();
  await ctx.reply("🚫 Session cancelled.");
}

export async function handleStatus(ctx: CommandContext): Promise<void> {
  if (ctx.sessionId) {
    const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
    if (session) {
      await ctx.reply(`📊 Session: ${session.name} (${session.status})`);
      return;
    }
  }
  await ctx.reply("📊 Status check (not yet listing all sessions)");
}

export async function handleSessions(ctx: CommandContext): Promise<void> {
  // TODO: Implement session listing once SessionManager API is confirmed
  await ctx.reply("📋 Sessions list not yet implemented");
}

export async function handleHandoff(ctx: CommandContext): Promise<void> {
  if (!ctx.sessionId) {
    await ctx.reply("❌ No active session.");
    return;
  }
  const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
  if (!session) {
    await ctx.reply("❌ Session not found.");
    return;
  }
  await ctx.reply(`🔄 Handoff not yet implemented for session ${ctx.sessionId}`);
}

export async function executeCancelSession(ctx: CommandContext): Promise<void> {
  await handleCancel(ctx);
}