import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";
import { sendCard } from "../send-utils.js";
import { buildLevel1, buildLevel2, escapeMd } from "../message-composer.js";

/** Send a one-shot info card matching the standard Container style. */
async function sendInfoCard(ctx: CommandContext, emoji: string, label: string, detail: string): Promise<void> {
  const card = {
    type: "AdaptiveCard",
    version: "1.4",
    body: [{
      type: "Container",
      spacing: "Small",
      items: [
        buildLevel1(emoji, escapeMd(label)),
        buildLevel2(detail),
      ],
    }],
    width: "stretch",
  };
  await sendCard(ctx.context, card as Record<string, unknown>);
}

/**
 * Handle /cancel — finalize the current card and destroy the session.
 */
export async function handleCancel(ctx: CommandContext): Promise<void> {
  if (!ctx.sessionId) {
    await sendInfoCard(ctx, "❌", "Error", "No active session to cancel.");
    return;
  }
  const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
  if (!session) {
    await sendInfoCard(ctx, "❌", "Error", "Session not found.");
    return;
  }

  try {
    await ctx.adapter["composer"].finalize(ctx.sessionId);
    try { await session.destroy(); } catch { /* best effort */ }
    await sendInfoCard(ctx, "🚫", "Cancelled", session.name || ctx.sessionId.slice(0, 8));
  } catch (err) {
    log.error({ err, sessionId: ctx.sessionId }, "[session] cancel failed");
    await sendInfoCard(ctx, "❌", "Failed", "Could not cancel session.");
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
        `**Session:** ${session.name || session.id.slice(0, 8)}\n\n` +
        `**Agent:** ${session.agentName}\n\n` +
        `**Status:** ${session.status}\n\n` +
        `**Workspace:** \`${session.workingDirectory}\`\n\n` +
        `**Queue:** ${session.queueDepth} pending`,
      );
      return;
    }

    // Fallback to stored record
    const record = ctx.adapter.core.sessionManager.getSessionRecord(ctx.sessionId);
    if (record) {
      await ctx.reply(
        `**Session:** ${record.name || record.sessionId.slice(0, 8)}\n\n` +
        `**Agent:** ${record.agentName}\n\n` +
        `**Status:** ${record.status} (not loaded)\n\n` +
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
    `**OpenACP Status**\n\n` +
    `Active sessions: ${active.length}\n\n` +
    `Error sessions: ${errors.length}\n\n` +
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

  const truncated = allRecords.length > 20 ? `\n\n...and ${allRecords.length - 20} more` : "";

  await ctx.reply(`**Sessions: ${allRecords.length}**\n\n${lines.join("\n\n")}${truncated}`);
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
