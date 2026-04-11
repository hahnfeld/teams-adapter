import type { CommandContext } from "./index.js";
import type { CommandRegistry } from "@openacp/plugin-sdk";

/**
 * Handle /doctor — run system diagnostics.
 * Delegates to the core doctor command if available, otherwise checks basics.
 */
export async function handleDoctor(ctx: CommandContext): Promise<void> {
  const registry = ctx.adapter.core.lifecycleManager?.serviceRegistry?.get<CommandRegistry>("command-registry");
  if (registry) {
    const def = registry.get("doctor");
    if (def) {
      try {
        const response = await registry.execute("/doctor", {
          raw: "",
          sessionId: null,
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

  // Basic self-diagnostics
  const checks: string[] = [];

  // Bot connectivity
  checks.push("✅ Bot is running and receiving messages");

  // Session manager
  const allRecords = ctx.adapter.core.sessionManager.listRecords();
  const active = allRecords.filter((r) => r.status === "active" || r.status === "initializing");
  const errors = allRecords.filter((r) => r.status === "error");
  checks.push(`✅ Session store: ${allRecords.length} sessions (${active.length} active, ${errors.length} errors)`);

  // Agent manager
  const agents = ctx.adapter.core.agentManager.getAvailableAgents();
  checks.push(`${agents.length > 0 ? "✅" : "⚠️"} Agents available: ${agents.length}`);

  // Graph API
  const hasGraph = ctx.adapter.getTeamId() && ctx.adapter.getChannelId();
  checks.push(`${hasGraph ? "✅" : "ℹ️"} Teams config: teamId=${ctx.adapter.getTeamId() ? "set" : "missing"}, channelId=${ctx.adapter.getChannelId() ? "set" : "missing"}`);

  // Assistant
  const assistantId = ctx.adapter.getAssistantSessionId();
  checks.push(`${assistantId ? "✅" : "ℹ️"} Assistant session: ${assistantId ? "active" : "not configured"}`);

  await ctx.reply(`**🔍 System Diagnostics**\n\n${checks.join("\n\n")}`);
}

export async function handleDoctorButton(ctx: CommandContext): Promise<void> {
  await handleDoctor(ctx);
}
