import type { CommandContext } from "./index.js";
import type { CommandRegistry } from "@openacp/plugin-sdk";

/**
 * Handle /settings — show current configuration.
 * Delegates to core settings command if available, otherwise shows adapter config.
 */
export async function handleSettings(ctx: CommandContext): Promise<void> {
  const registry = ctx.adapter.core.lifecycleManager?.serviceRegistry?.get<CommandRegistry>("command-registry");
  if (registry) {
    const def = registry.get("settings");
    if (def) {
      try {
        const response = await registry.execute("/settings", {
          raw: "",
          sessionId: ctx.sessionId,
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
  }

  // Show adapter-level config
  const config = ctx.adapter.core.configManager.get();
  const defaultAgent = config.defaultAgent ?? "not set";
  const workspace = ctx.adapter.core.configManager.resolveWorkspace?.() ?? "not set";

  let sessionInfo = "";
  if (ctx.sessionId) {
    const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
    if (session) {
      sessionInfo = `\n\n**Session Settings:**\n` +
        `- Agent: ${session.agentName}\n` +
        `- Mode: ${session.getConfigByCategory?.("mode")?.currentValue ?? "default"}\n` +
        `- Model: ${session.getConfigByCategory?.("model")?.currentValue ?? "default"}\n` +
        `- Bypass: ${session.clientOverrides?.bypassPermissions ? "on" : "off"}\n` +
        `- TTS: ${session.voiceMode ?? "off"}`;
    }
  }

  await ctx.reply(
    `**⚙️ Configuration**\n\n` +
    `**Global:**\n` +
    `- Default agent: ${defaultAgent}\n` +
    `- Workspace: \`${workspace}\`\n` +
    `- Teams channel: ${ctx.adapter.getChannelId() || "not set"}\n` +
    `- Notification channel: ${ctx.adapter.getAssistantThreadId() ? "configured" : "not set"}\n` +
    `- Graph API: ${ctx.adapter.getTeamId() ? "configured" : "not configured"}` +
    sessionInfo,
  );
}

export async function handleSettingsButton(ctx: CommandContext): Promise<void> {
  await handleSettings(ctx);
}
