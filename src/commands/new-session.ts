import { sendInfoCard } from "./index.js";
import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";
import { sendCard } from "../send-utils.js";
import { buildLevel1 } from "../message-composer.js";

/**
 * Handle /new [agent] [workspace] — create a new agent session.
 * Mirrors Telegram's createSessionDirect pattern.
 */
export async function handleNew(ctx: CommandContext, args: string[]): Promise<void> {
  const agentName = args[0];
  const workspace = args[1];

  if (!agentName) {
    // Don't destroy the existing session here — the user hasn't committed yet.
    // The session is destroyed in handleDialogAction("new-session") when the
    // user actually submits the wizard form.

    // Send the session wizard inline as an Adaptive Card.
    // Uses the same Container + ColumnSet pattern as all other card entries.
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
        {
          type: "Container",
          spacing: "Small",
          items: [
            buildLevel1("➕", "New Session"),
            {
              type: "ColumnSet",
              spacing: "None",
              columns: [
                { type: "Column", width: "20px" },
                {
                  type: "Column",
                  width: "auto",
                  items: [{ type: "TextBlock", text: "⎿", size: "Small", fontType: "Monospace", spacing: "None" }],
                  verticalContentAlignment: "Top",
                },
                {
                  type: "Column",
                  width: "stretch",
                  items: [
                    { type: "TextBlock", text: "Agent", size: "Small", fontType: "Monospace", spacing: "None" },
                    { type: "Input.ChoiceSet", id: "agent", style: "compact", value: agentChoices[0].value, choices: agentChoices, spacing: "None" },
                    { type: "TextBlock", text: "Workspace", size: "Small", fontType: "Monospace", spacing: "Small" },
                    { type: "Input.Text", id: "workspace", placeholder: "/path/to/project", value: defaultWorkspace, spacing: "None" },
                  ],
                },
              ],
            },
            {
              type: "ActionSet",
              spacing: "Small",
              actions: [
                { type: "Action.Execute", title: "Create Session", verb: "dialog:new-session" },
              ],
            },
          ],
        },
      ],
      msteams: { width: "Full" },
    };
    await sendCard(ctx.context, card as Record<string, unknown>);
    return;
  }

  const workDir = workspace ?? ctx.adapter.core.configManager.resolveWorkspace?.() ?? process.cwd();

  try {
    // Cancel the old session before creating a new one
    if (ctx.sessionId) {
      await ctx.adapter["composer"].finalize(ctx.sessionId);
      const oldSession = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
      if (oldSession) {
        try { await oldSession.destroy(); } catch { /* best effort */ }
      }
    }

    await sendInfoCard(ctx, "🔧", "Creating session", agentName);

    const conversationId = (ctx.context.activity?.conversation?.id as string | undefined)?.split(";")[0];
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

    await sendInfoCard(ctx, "✅", "Session created", `${agentName} · ${workDir}`);
  } catch (err) {
    log.error({ err, agentName }, "[new-session] Failed to create session");
    await sendInfoCard(ctx, "❌", "Failed", err instanceof Error ? err.message : String(err));
  }
}

/**
 * Handle /newchat — start a new chat with the same agent and workspace
 * as the current session. Mirrors Telegram's "same context, fresh start" pattern.
 */
export async function handleNewChat(ctx: CommandContext): Promise<void> {
  if (!ctx.sessionId) {
    await sendInfoCard(ctx, "❌", "Error", "No active session. Use /new to create one.");
    return;
  }

  const session = ctx.adapter.core.sessionManager.getSession(ctx.sessionId);
  if (!session) {
    await sendInfoCard(ctx, "❌", "Error", "Session not found.");
    return;
  }

  const agentName = session.agentName;
  const workspace = session.workingDirectory;

  try {
    // Cancel the old session before creating a new one
    await ctx.adapter["composer"].finalize(ctx.sessionId);
    try { await session.destroy(); } catch { /* best effort */ }

    await sendInfoCard(ctx, "🔧", "Starting new chat", agentName);

    const conversationId = (ctx.context.activity?.conversation?.id as string | undefined)?.split(";")[0];
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

    await sendInfoCard(ctx, "✅", "New chat started", `${agentName} · ${workspace}`);
  } catch (err) {
    log.error({ err }, "[new-session] Failed to create new chat");
    await sendInfoCard(ctx, "❌", "Failed", err instanceof Error ? err.message : String(err));
  }
}
