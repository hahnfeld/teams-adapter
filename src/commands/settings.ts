import { sendInfoCard } from "./index.js";
import type { CommandContext } from "./index.js";
import type { CommandRegistry } from "@openacp/plugin-sdk";

/**
 * Handle /settings — show current configuration.
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
          reply: async (content: string) => {
            await sendInfoCard(ctx, "⚙️", "Settings", content);
          },
        });
        if (response.type === "text") {
          await sendInfoCard(ctx, "⚙️", "Settings", response.text);
        }
        return;
      } catch { /* fall through */ }
    }
  }

  const config = ctx.adapter.core.configManager.get();
  const defaultAgent = config.defaultAgent ?? "not set";
  const workspace = ctx.adapter.core.configManager.resolveWorkspace?.() ?? "not set";

  const lines: string[] = [
    `Default agent: ${defaultAgent}`,
    `Workspace: ${workspace}`,
    `Channel: ${ctx.adapter.getChannelId() || "not set"}`,
    `Graph API: ${ctx.adapter.getTeamId() ? "configured" : "not configured"}`,
  ];

  if (ctx.sessionId) {
    const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
    if (session) {
      lines.push("");
      lines.push(`Session: ${session.agentName}`);
      lines.push(`Mode: ${session.getConfigByCategory?.("mode")?.currentValue ?? "default"}`);
      lines.push(`Model: ${session.getConfigByCategory?.("model")?.currentValue ?? "default"}`);
      lines.push(`Bypass: ${session.clientOverrides?.bypassPermissions ? "on" : "off"}`);
      lines.push(`TTS: ${session.voiceMode ?? "off"}`);
    }
  }

  await sendInfoCard(ctx, "⚙️", "Settings", lines.join("\n"));
}
