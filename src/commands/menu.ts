import type { CommandContext } from "./index.js";

export async function handleMenu(ctx: CommandContext): Promise<void> {
  await ctx.reply("📋 Menu not yet implemented");
}

export async function handleHelp(ctx: CommandContext): Promise<void> {
  const commands = [
    "/new [agent] - Create new session",
    "/newchat - New chat, same agent & workspace",
    "/cancel - Cancel current session",
    "/status - Show session status",
    "/sessions - List all sessions",
    "/agents - List available agents",
    "/menu - Show action menu",
    "/help - Show this help",
    "/outputmode low|medium|high - Set output detail",
    "/bypass - Auto-approve permissions",
  ];
  await ctx.reply(`**Commands:**\n${commands.join("\n")}`);
}

export async function handleClear(ctx: CommandContext): Promise<void> {
  await ctx.reply("🗑️ Clear not yet implemented");
}

export async function handleMenuButton(ctx: CommandContext): Promise<void> {
  await handleMenu(ctx);
}