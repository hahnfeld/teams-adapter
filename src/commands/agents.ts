import type { CommandContext } from "./index.js";

export async function handleAgents(ctx: CommandContext): Promise<void> {
  await ctx.reply("🧠 Agents list not yet implemented");
}

export async function handleInstall(ctx: CommandContext, name?: string): Promise<void> {
  if (!name) {
    await ctx.reply("⚠️ Usage: /install <agent-name>");
    return;
  }
  await ctx.reply(`📦 Installing ${name}... (not yet implemented)`);
}

export async function handleAgentButton(ctx: CommandContext): Promise<void> {
  await ctx.reply("Agent action not yet implemented");
}