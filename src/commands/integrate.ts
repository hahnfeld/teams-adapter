import type { CommandContext } from "./index.js";

export async function handleIntegrate(ctx: CommandContext): Promise<void> {
  await ctx.reply("🔗 Integration management not yet implemented");
}

export async function handleIntegrateButton(ctx: CommandContext): Promise<void> {
  await handleIntegrate(ctx);
}