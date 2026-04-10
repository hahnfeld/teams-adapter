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
  const sessions = ctx.adapter.core.sessionManager.getAllSessions();
  const active = sessions.filter((s) => s.status === "active");

  if (ctx.sessionId) {
    const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
    if (session) {
      await ctx.reply(`📊 Session: ${session.name} (${session.status})`);
      return;
    }
  }

  await ctx.reply(`📊 ${active.length} active session(s) of ${sessions.length} total`);
}

export async function handleSessions(ctx: CommandContext): Promise<void> {
  const sessions = ctx.adapter.core.sessionManager.getAllSessions();
  if (sessions.length === 0) {
    await ctx.reply("No sessions.");
    return;
  }
  const lines = sessions.slice(0, 10).map((s) => `• **${s.name}** (${s.status})`).join("\n");
  await ctx.reply(`Sessions:\n${lines}`);
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
  // TODO: Generate terminal resume command
  await ctx.reply(`🔄 Handoff not yet implemented for session ${ctx.sessionId}`);
}

export async function executeCancelSession(ctx: CommandContext): Promise<void> {
  await handleCancel(ctx);
}