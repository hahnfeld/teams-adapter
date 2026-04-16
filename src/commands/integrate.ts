import { sendInfoCard } from "./index.js";
import type { CommandContext } from "./index.js";
import type { CommandRegistry } from "@openacp/plugin-sdk";

/**
 * Handle /integrate — manage agent integrations.
 */
export async function handleIntegrate(ctx: CommandContext): Promise<void> {
  const registry = ctx.adapter.core.lifecycleManager?.serviceRegistry?.get<CommandRegistry>("command-registry");
  if (registry) {
    const def = registry.get("integrate");
    if (def) {
      try {
        const response = await registry.execute("/integrate", {
          raw: "",
          sessionId: ctx.sessionId,
          channelId: "teams",
          userId: ctx.userId,
          reply: async (content: string) => {
            await sendInfoCard(ctx, "🔗", "Integrate", content);
          },
        });
        if (response.type !== "silent") {
          if (response.type === "text") {
            await sendInfoCard(ctx, "🔗", "Integrate", response.text);
          } else if (response.type === "list") {
            const items = response.items.map((i: any) => `- ${i.label}${i.detail ? ` — ${i.detail}` : ""}`).join("\n");
            await sendInfoCard(ctx, "🔗", response.title, items);
          }
        }
        return;
      } catch { /* fall through */ }
    }
  }

  const agents = ctx.adapter.core.agentManager.getAvailableAgents();
  if (agents.length === 0) {
    await sendInfoCard(ctx, "🔗", "Integrate", "No agents installed. Use /agents to browse.");
    return;
  }

  const lines = agents.map((a) => `- ${a.name}`);
  await sendInfoCard(ctx, "🔗", "Integrations", `${lines.join("\n")}\n\nUse openacp integrate from terminal for details.`);
}
