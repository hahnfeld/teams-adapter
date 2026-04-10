import type { CommandContext } from "./index.js";
import type { CommandRegistry } from "@openacp/plugin-sdk";

/**
 * Handle /integrate — manage agent integrations (handoff, tools, etc.).
 * Delegates to core integrate command if available.
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
          reply: async (content: string) => { await ctx.reply(content); },
        });
        if (response.type !== "silent") {
          if (response.type === "text") await ctx.reply(response.text);
          else if (response.type === "list") {
            const items = response.items.map((i: any) => `- **${i.label}**${i.detail ? ` — ${i.detail}` : ""}`).join("\n");
            await ctx.reply(`${response.title}\n${items}`);
          }
        }
        return;
      } catch { /* fall through */ }
    }
  }

  // Show available integrations
  const agents = ctx.adapter.core.agentManager.getAvailableAgents();
  if (agents.length === 0) {
    await ctx.reply("No agents installed. Use `/agents` to browse and install.");
    return;
  }

  const lines = agents.map((a) => `- **${a.name}**`);
  await ctx.reply(
    `**🔗 Integrations**\n\n` +
    `Installed agents:\n${lines.join("\n")}\n\n` +
    `Use \`openacp integrate <agent>\` from the terminal for detailed integration management.`,
  );
}

export async function handleIntegrateButton(ctx: CommandContext): Promise<void> {
  await handleIntegrate(ctx);
}
