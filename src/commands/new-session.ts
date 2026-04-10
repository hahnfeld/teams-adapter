import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";

/**
 * Handle /new [agent] [workspace] — create a new agent session.
 * Mirrors Telegram's createSessionDirect pattern.
 */
export async function handleNew(ctx: CommandContext, args: string[]): Promise<void> {
  const agentName = args[0];
  const workspace = args[1];

  if (!agentName) {
    const agents = ctx.adapter.core.agentManager.getAvailableAgents();
    const agentList = agents.map((a) => `- ${a.name}`).join("\n");
    await ctx.reply(
      `**Create a new session**\n\n` +
      `Usage: \`/new <agent> [workspace]\`\n\n` +
      `**Available agents:**\n${agentList || "_No agents installed_"}`,
    );
    return;
  }

  const workDir = workspace ?? ctx.adapter.core.configManager.resolveWorkspace?.() ?? process.cwd();

  try {
    await ctx.reply(`🔄 Creating session with **${agentName}**...`);

    const session = await ctx.adapter.core.sessionManager.createSession(
      "teams",
      agentName,
      workDir,
      ctx.adapter.core.agentManager,
    );

    const threadId = await ctx.adapter.createSessionThread(session.id, session.name || agentName);
    session.threadId = threadId;
    session.threadIds.set("teams", threadId);

    await ctx.reply(
      `✅ Session created\n` +
      `**Agent:** ${agentName}\n` +
      `**Workspace:** \`${workDir}\`\n` +
      `**Session:** ${session.id.slice(0, 8)}`,
    );
  } catch (err) {
    log.error({ err, agentName }, "[new-session] Failed to create session");
    await ctx.reply(`❌ Failed to create session: ${err instanceof Error ? err.message : String(err)}`);
  }
}

/**
 * Handle /newchat — start a new chat with the same agent and workspace
 * as the current session. Mirrors Telegram's "same context, fresh start" pattern.
 */
export async function handleNewChat(ctx: CommandContext): Promise<void> {
  if (!ctx.sessionId) {
    await ctx.reply("❌ No active session. Use `/new <agent>` to create one.");
    return;
  }

  const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
  if (!session) {
    await ctx.reply("❌ Session not found.");
    return;
  }

  const agentName = session.agentName;
  const workspace = session.workingDirectory;

  try {
    await ctx.reply(`🔄 Starting new chat with **${agentName}**...`);

    const newSession = await ctx.adapter.core.sessionManager.createSession(
      "teams",
      agentName,
      workspace,
      ctx.adapter.core.agentManager,
    );

    const threadId = await ctx.adapter.createSessionThread(newSession.id, newSession.name || agentName);
    newSession.threadId = threadId;
    newSession.threadIds.set("teams", threadId);

    await ctx.reply(
      `✅ New chat started\n` +
      `**Agent:** ${agentName}\n` +
      `**Session:** ${newSession.id.slice(0, 8)}`,
    );
  } catch (err) {
    log.error({ err }, "[new-session] Failed to create new chat");
    await ctx.reply(`❌ Failed: ${err instanceof Error ? err.message : String(err)}`);
  }
}

export async function executeNewSession(
  ctx: CommandContext,
  agentName?: string,
  workspace?: string,
): Promise<void> {
  await handleNew(ctx, [agentName ?? "", workspace ?? ""].filter(Boolean));
}
