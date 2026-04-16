import { sendInfoCard } from "./index.js";
import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";
import type { CommandRegistry } from "@openacp/plugin-sdk";

/**
 * Handle /bypass — toggle auto-approve permissions for the current session.
 */
export async function handleBypass(ctx: CommandContext): Promise<void> {
  if (!ctx.sessionId) {
    await sendInfoCard(ctx, "❌", "Error", "No active session. Bypass applies per-session.");
    return;
  }

  const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
  if (!session) {
    await sendInfoCard(ctx, "❌", "Error", "Session not found.");
    return;
  }

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
      const emoji = newState ? "☠️" : "🔐";
      const label = newState ? "Bypass enabled" : "Bypass disabled";
      const detail = newState ? "Permissions auto-approved" : "Approvals required";
      await sendInfoCard(ctx, emoji, label, detail);
      return;
    } catch { /* fall through */ }
  }

  const currentBypass = !!session.clientOverrides?.bypassPermissions;
  session.clientOverrides = { ...session.clientOverrides, bypassPermissions: !currentBypass };
  const emoji = !currentBypass ? "☠️" : "🔐";
  const label = !currentBypass ? "Bypass enabled" : "Bypass disabled";
  const detail = !currentBypass ? "Permissions auto-approved" : "Approvals required";
  await sendInfoCard(ctx, emoji, label, detail);
}

/**
 * Handle /tts [on|off] — toggle text-to-speech for the current session.
 */
export async function handleTTS(ctx: CommandContext, mode?: string): Promise<void> {
  if (!ctx.sessionId) {
    await sendInfoCard(ctx, "❌", "Error", "No active session.");
    return;
  }

  const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
  if (!session) {
    await sendInfoCard(ctx, "❌", "Error", "Session not found.");
    return;
  }

  if (mode === "on" || mode === "off") {
    session.voiceMode = mode;
    await sendInfoCard(ctx, "🔊", "TTS", mode === "on" ? "Enabled" : "Disabled");
    return;
  }

  const newMode = session.voiceMode === "on" ? "off" : "on";
  session.voiceMode = newMode;
  await sendInfoCard(ctx, "🔊", "TTS", newMode === "on" ? "Enabled" : "Disabled");
}

/**
 * Handle /restart — restart the OpenACP assistant session.
 */
export async function handleRestart(ctx: CommandContext): Promise<void> {
  try {
    await ctx.adapter.restartAssistant();
    await sendInfoCard(ctx, "🔄", "Restarting", "OpenACP assistant");
  } catch (err) {
    log.error({ err }, "[admin] restartAssistant failed");
    await sendInfoCard(ctx, "❌", "Restart failed", "Check logs for details.");
  }
}

/**
 * Handle /respawn — restart the assistant session.
 */
export async function handleRespawn(ctx: CommandContext): Promise<void> {
  try {
    await ctx.adapter.respawnAssistant();
    await sendInfoCard(ctx, "🔄", "Respawned", "Assistant session restarted.");
  } catch (err) {
    log.error({ err }, "[admin] respawnAssistant failed");
    await sendInfoCard(ctx, "❌", "Respawn failed", "Check logs for details.");
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
        reply: async (content: string) => {
          await sendInfoCard(ctx, "📦", "Update", content);
        },
      });
      if (response.type === "text") {
        await sendInfoCard(ctx, "📦", "Update", response.text);
      }
      return;
    } catch { /* fall through */ }
  }
  await sendInfoCard(ctx, "📦", "Update", "Not available. Run openacp update from the terminal.");
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
    await sendInfoCard(ctx, "🔄", "Output mode", "Reset to medium");
    return;
  }

  if (level !== "low" && level !== "medium" && level !== "high") {
    await sendInfoCard(ctx, "⚠️", "Output mode", "Valid levels: low, medium, high, reset");
    return;
  }

  if (scope === "session" && ctx.sessionId) {
    try {
      await ctx.adapter.core.sessionManager.patchRecord(ctx.sessionId, { outputMode: level } as any);
    } catch { /* best effort */ }
    ctx.adapter.setSessionOutputMode(ctx.sessionId, level as "low" | "medium" | "high");
    await sendInfoCard(ctx, "🔄", "Output mode", `${level} (this session)`);
  } else if (ctx.sessionId) {
    ctx.adapter.setSessionOutputMode(ctx.sessionId, level as "low" | "medium" | "high");
    await sendInfoCard(ctx, "🔄", "Output mode", `${level} (this session)`);
  } else {
    await sendInfoCard(ctx, "🔄", "Output mode", `${level} (use in a session for per-session control)`);
  }
}
