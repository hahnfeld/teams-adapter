import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";
import { sendCard } from "../send-utils.js";

/**
 * Handle /new [agent] [workspace] — create a new agent session.
 * Mirrors Telegram's createSessionDirect pattern.
 */
export async function handleNew(ctx: CommandContext, args: string[]): Promise<void> {
  const agentName = args[0];
  const workspace = args[1];

  if (!agentName) {
    // Send the session wizard inline as an Adaptive Card.
    // Task modules (task/fetch popups) don't work reliably for sideloaded apps,
    // so we render the wizard directly in the chat.
    const agents = ctx.adapter.core.agentManager.getAvailableAgents();
    const defaultWorkspace = ctx.adapter.core.configManager.resolveWorkspace?.() ?? process.cwd();
    const agentChoices = agents.map((a: { name: string }) => ({ title: a.name, value: a.name }));
    if (agentChoices.length === 0) {
      agentChoices.push({ title: "openacp", value: "openacp" });
    }
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        { type: "TextBlock", text: "**New Session**", weight: "Bolder", size: "Large" },
        { type: "TextBlock", text: "Select an agent and workspace to start a coding session.", wrap: true, isSubtle: true, spacing: "Small" },
        { type: "TextBlock", text: "Agent", weight: "Bolder", spacing: "Large" },
        { type: "Input.ChoiceSet", id: "agent", style: "compact", value: agentChoices[0].value, choices: agentChoices },
        { type: "TextBlock", text: "Workspace (project directory)", weight: "Bolder", spacing: "Large" },
        { type: "Input.Text", id: "workspace", placeholder: "/path/to/project", value: defaultWorkspace },
      ],
      actions: [
        { type: "Action.Execute", title: "Create Session", verb: "dialog:new-session" },
      ],
    };
    await sendCard(ctx.context, card as Record<string, unknown>);
    return;
  }

  const workDir = workspace ?? ctx.adapter.core.configManager.resolveWorkspace?.() ?? process.cwd();

  try {
    await ctx.reply(`🔄 Creating session with **${agentName}**...`);

    const conversationId = ctx.context.activity?.conversation?.id as string | undefined;
    const session = await (ctx.adapter.core as any).createSession({
      channelId: "teams",
      agentName,
      workingDirectory: workDir,
      threadId: conversationId,
      createThread: !conversationId,
    });
    if (conversationId) {
      session.threadIds.set("teams", conversationId);
    }
    ctx.adapter["_sessionContexts"].set(session.id, { context: ctx.context, isAssistant: false, threadId: conversationId });

    await ctx.reply(
      `✅ Session created\n\n` +
      `**Agent:** ${agentName}\n\n` +
      `**Workspace:** \`${workDir}\`\n\n` +
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

    const conversationId = ctx.context.activity?.conversation?.id as string | undefined;
    const newSession = await (ctx.adapter.core as any).createSession({
      channelId: "teams",
      agentName,
      workingDirectory: workspace,
      threadId: conversationId,
      createThread: !conversationId,
    });
    if (conversationId) {
      newSession.threadIds.set("teams", conversationId);
    }
    ctx.adapter["_sessionContexts"].set(newSession.id, { context: ctx.context, isAssistant: false, threadId: conversationId });

    await ctx.reply(
      `✅ New chat started\n\n` +
      `**Agent:** ${agentName}\n\n` +
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
