import { sendInfoCard } from "./index.js";
import type { CommandContext } from "./index.js";
import type { CommandRegistry } from "@openacp/plugin-sdk";

/**
 * Handle /doctor — run system diagnostics.
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
          reply: async (content: string) => {
            await sendInfoCard(ctx, "🔍", "Doctor", content);
          },
        });
        if (response.type === "text") {
          await sendInfoCard(ctx, "🔍", "Doctor", response.text);
        }
        return;
      } catch { /* fall through */ }
    }
  }

  // Basic self-diagnostics
  const checks: string[] = [];
  checks.push("✓ Bot running");

  const allRecords = ctx.adapter.core.sessionManager.listRecords();
  const active = allRecords.filter((r) => r.status === "active" || r.status === "initializing");
  const errors = allRecords.filter((r) => r.status === "error");
  checks.push(`✓ Sessions: ${allRecords.length} total, ${active.length} active, ${errors.length} errors`);

  const agents = ctx.adapter.core.agentManager.getAvailableAgents();
  checks.push(`${agents.length > 0 ? "✓" : "⚠"} Agents: ${agents.length}`);

  const hasGraph = ctx.adapter.getTeamId() && ctx.adapter.getChannelId();
  checks.push(`${hasGraph ? "✓" : "⚠"} Teams: teamId=${ctx.adapter.getTeamId() ? "set" : "missing"}, channelId=${ctx.adapter.getChannelId() ? "set" : "missing"}`);

  const assistantId = ctx.adapter.getAssistantSessionId();
  checks.push(`${assistantId ? "✓" : "◻"} Assistant: ${assistantId ? "active" : "not configured"}`);

  await sendInfoCard(ctx, "🔍", "Doctor", checks.join("\n"));
}
