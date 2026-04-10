import type { CommandContext } from "./index.js";

export async function handleDoctor(ctx: CommandContext): Promise<void> {
  await ctx.reply("🔍 Running diagnostics... (not yet implemented)");
}

export async function handleDoctorButton(ctx: CommandContext): Promise<void> {
  await handleDoctor(ctx);
}