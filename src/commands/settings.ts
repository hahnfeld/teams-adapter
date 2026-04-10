import type { CommandContext } from "./index.js";

export async function handleSettings(ctx: CommandContext): Promise<void> {
  await ctx.reply("⚙️ Settings not yet implemented");
}

export async function handleSettingsButton(ctx: CommandContext): Promise<void> {
  await handleSettings(ctx);
}