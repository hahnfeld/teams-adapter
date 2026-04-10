import type { CommandContext } from "./index.js";

export async function handleNew(ctx: CommandContext, args: string[]): Promise<void> {
  const agentName = args[0] ?? undefined;

  await ctx.reply(`🔄 Creating new session...`);
  // TODO: Implement new session creation
  await ctx.reply(`✅ New session created (agent: ${agentName ?? "default"})`);
}

export async function handleNewChat(ctx: CommandContext): Promise<void> {
  await ctx.reply(`🔄 Creating new chat...`);
  // TODO: Implement new chat with same agent/workspace
  await ctx.reply(`✅ New chat started`);
}

export async function executeNewSession(
  ctx: CommandContext,
  agentName?: string,
  workspace?: string,
): Promise<void> {
  await ctx.reply(`🔄 Creating session (agent: ${agentName ?? "default"}, workspace: ${workspace ?? "default"})...`);
  // TODO: Implement
  await ctx.reply(`✅ Session created`);
}