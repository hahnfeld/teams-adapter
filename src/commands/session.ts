import { sendInfoCard } from "./index.js";
import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";

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
      await sendInfoCard(ctx, "📊", "Status",
        `${session.name || session.id.slice(0, 8)}\n` +
        `Agent: ${session.agentName}\n` +
        `Status: ${session.status}\n` +
        `Workspace: ${session.workingDirectory}\n` +
        `Queue: ${session.queueDepth} pending`,
      );
      return;
    }

    const record = ctx.adapter.core.sessionManager.getSessionRecord(ctx.sessionId);
    if (record) {
      await sendInfoCard(ctx, "📊", "Status",
        `${record.name || record.sessionId.slice(0, 8)}\n` +
        `Agent: ${record.agentName}\n` +
        `Status: ${record.status} (not loaded)\n` +
        `Workspace: ${record.workingDir}`,
      );
      return;
    }
  }

  const allRecords = ctx.adapter.core.sessionManager.listRecords();
  const active = allRecords.filter((r) => r.status === "active" || r.status === "initializing");
  const errors = allRecords.filter((r) => r.status === "error");

  await sendInfoCard(ctx, "📊", "Status",
    `Active: ${active.length}\n` +
    `Errors: ${errors.length}\n` +
    `Total: ${allRecords.length}`,
  );
}

/**
 * Handle /sessions — list all sessions with status.
 * Mirrors Telegram's handleTopics with status overview.
 */
export async function handleSessions(ctx: CommandContext): Promise<void> {
  const allRecords = ctx.adapter.core.sessionManager.listRecords();

  if (allRecords.length === 0) {
    await sendInfoCard(ctx, "📋", "Sessions", "None. Use /new to create one.");
    return;
  }

  const statusEmoji: Record<string, string> = {
    active: "🟢", initializing: "🟡", finished: "✓", error: "❌", cancelled: "⛔",
  };
  const statusOrder: Record<string, number> = { active: 0, initializing: 1, error: 2, finished: 3, cancelled: 4 };
  allRecords.sort((a, b) => (statusOrder[a.status] ?? 5) - (statusOrder[b.status] ?? 5));

  const lines = allRecords.slice(0, 20).map((r) => {
    const emoji = statusEmoji[r.status] || "⚪";
    const name = r.name?.trim() || `${r.agentName} session`;
    return `${emoji} ${name} — ${r.status}`;
  });
  const truncated = allRecords.length > 20 ? `\n...and ${allRecords.length - 20} more` : "";

  await sendInfoCard(ctx, "📋", `Sessions (${allRecords.length})`, lines.join("\n") + truncated);
}

/**
 * Handle /handoff — generate a terminal resume command for the current session.
 */
export async function handleHandoff(ctx: CommandContext): Promise<void> {
  if (!ctx.sessionId) {
    await sendInfoCard(ctx, "❌", "Error", "No active session.");
    return;
  }
  const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
  if (!session) {
    await sendInfoCard(ctx, "❌", "Error", "Session not found.");
    return;
  }
  await sendInfoCard(ctx, "🔗", "Handoff", `openacp adopt ${ctx.sessionId}`);
}
