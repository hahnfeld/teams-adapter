import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";

/**
 * Handle /cancel — abort the current prompt in the active session.
 * Mirrors Telegram's handleCancel pattern: abortPrompt, not destroy.
 */
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

  try {
    await session.abortPrompt();
    await ctx.reply("⛔ Prompt aborted. Session is still active — send a new message to continue.");
  } catch (err) {
    log.error({ err, sessionId: ctx.sessionId }, "[session] abortPrompt failed");
    // Fallback: destroy the session
    try {
      await session.destroy();
      await ctx.reply("🚫 Session cancelled.");
    } catch {
      await ctx.reply("❌ Failed to cancel session.");
    }
  }
}

/**
 * Handle /status — show session info or overall system stats.
 * Mirrors Telegram's handleStatus with both session-level and global views.
 */
export async function handleStatus(ctx: CommandContext): Promise<void> {
  if (ctx.sessionId) {
    const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
    if (session) {
      await ctx.reply(
        `**Session:** ${session.name || session.id.slice(0, 8)}\n` +
        `**Agent:** ${session.agentName}\n` +
        `**Status:** ${session.status}\n` +
        `**Workspace:** \`${session.workingDirectory}\`\n` +
        `**Queue:** ${session.queueDepth} pending`,
      );
      return;
    }

    // Fallback to stored record
    const record = ctx.adapter.core.sessionManager.getSessionRecord(ctx.sessionId);
    if (record) {
      await ctx.reply(
        `**Session:** ${record.name || record.sessionId.slice(0, 8)}\n` +
        `**Agent:** ${record.agentName}\n` +
        `**Status:** ${record.status} (not loaded)\n` +
        `**Workspace:** \`${record.workingDir}\``,
      );
      return;
    }
  }

  // Global status
  const allRecords = ctx.adapter.core.sessionManager.listRecords();
  const active = allRecords.filter((r) => r.status === "active" || r.status === "initializing");
  const errors = allRecords.filter((r) => r.status === "error");

  await ctx.reply(
    `**OpenACP Status**\n` +
    `Active sessions: ${active.length}\n` +
    `Error sessions: ${errors.length}\n` +
    `Total sessions: ${allRecords.length}`,
  );
}

/**
 * Handle /sessions — list all sessions with status.
 * Mirrors Telegram's handleTopics with status overview.
 */
export async function handleSessions(ctx: CommandContext): Promise<void> {
  const allRecords = ctx.adapter.core.sessionManager.listRecords();

  if (allRecords.length === 0) {
    await ctx.reply("No sessions found. Use `/new <agent>` to create one.");
    return;
  }

  const statusEmoji: Record<string, string> = {
    active: "🟢", initializing: "🟡", finished: "✅", error: "❌", cancelled: "⛔",
  };

  // Sort: active first
  const statusOrder: Record<string, number> = { active: 0, initializing: 1, error: 2, finished: 3, cancelled: 4 };
  allRecords.sort((a, b) => (statusOrder[a.status] ?? 5) - (statusOrder[b.status] ?? 5));

  const lines = allRecords.slice(0, 20).map((r) => {
    const emoji = statusEmoji[r.status] || "⚪";
    const name = r.name?.trim() || `${r.agentName} session`;
    return `${emoji} **${name}** — ${r.status}`;
  });

  const truncated = allRecords.length > 20 ? `\n\n_...and ${allRecords.length - 20} more_` : "";

  await ctx.reply(`**Sessions: ${allRecords.length}**\n\n${lines.join("\n")}${truncated}`);
}

/**
 * Handle /handoff — generate a terminal resume command for the current session.
 */
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
  await ctx.reply(
    `**Handoff to terminal:**\n\n` +
    `\`openacp adopt ${ctx.sessionId}\`\n\n` +
    `Run this command in a terminal to take over this session.`,
  );
}

export async function executeCancelSession(ctx: CommandContext): Promise<void> {
  await handleCancel(ctx);
}
