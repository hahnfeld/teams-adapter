import type { CommandContext } from "./index.js";
import { sendCard, sendActivity } from "../send-utils.js";

/**
 * Handle /menu — show the main action menu as an Adaptive Card with buttons.
 * Mirrors Telegram's buildMenuKeyboard pattern.
 */
export async function handleMenu(ctx: CommandContext): Promise<void> {
  const card = {
    type: "AdaptiveCard",
    version: "1.2",
    body: [
      { type: "TextBlock", text: "📋 **OpenACP Menu**", weight: "Bolder", size: "Medium" },
      { type: "TextBlock", text: "Quick actions for managing sessions and agents.", wrap: true, isSubtle: true },
    ],
    actions: [
      // Task Module dialog — opens a modal form for session creation
      {
        type: "Action.Submit",
        title: "➕ New Session",
        data: { msteams: { type: "task/fetch" }, dialogId: "new-session" },
      },
      // Task Module dialog — opens settings modal
      {
        type: "Action.Submit",
        title: "⚙️ Settings",
        data: { msteams: { type: "task/fetch" }, dialogId: "settings", sessionId: ctx.sessionId },
      },
      // Inline commands
      { type: "Action.Submit", title: "📊 Status", data: { verb: "cmd:status" } },
      { type: "Action.Submit", title: "📋 Sessions", data: { verb: "cmd:sessions" } },
      { type: "Action.Submit", title: "🔍 Doctor", data: { verb: "cmd:doctor" } },
    ],
  };

  await sendCard(ctx.context, card as Record<string, unknown>);
}

/**
 * Handle /help — show available commands.
 */
export async function handleHelp(ctx: CommandContext): Promise<void> {
  const commands = [
    "**Session Management:**",
    "`/new [agent] [workspace]` — Create new session",
    "`/newchat` — New chat, same agent & workspace",
    "`/cancel` — Abort current prompt",
    "`/status` — Show session status",
    "`/sessions` — List all sessions",
    "`/handoff` — Generate terminal resume command",
    "",
    "**Agent Management:**",
    "`/agents` — List available agents",
    "`/install <name>` — Install an agent",
    "",
    "**Settings:**",
    "`/outputmode low|medium|high` — Set output detail",
    "`/bypass` — Toggle auto-approve permissions",
    "`/tts on|off` — Toggle text-to-speech",
    "`/mode <mode>` — Switch session mode",
    "`/model <model>` — Switch AI model",
    "`/thought <level>` — Adjust thinking level",
    "`/settings` — Show configuration",
    "",
    "**System:**",
    "`/menu` — Show action menu",
    "`/doctor` — Run diagnostics",
    "`/restart` — Restart OpenACP",
    "`/clear` — Reset assistant session",
  ];
  // Send with suggested action buttons for quick access (1:1 chat only)
  await sendActivity(ctx.context, {
    text: commands.join("\n"),
    suggestedActions: {
      actions: [
        { type: "imBack", title: "➕ New", value: "/new" },
        { type: "imBack", title: "📊 Status", value: "/status" },
        { type: "imBack", title: "🤖 Agents", value: "/agents" },
        { type: "imBack", title: "📋 Menu", value: "/menu" },
      ],
    },
  });
}

/**
 * Handle /clear — reset the assistant session.
 */
export async function handleClear(ctx: CommandContext): Promise<void> {
  try {
    await ctx.adapter.respawnAssistant();
    await ctx.reply("🗑️ Assistant session cleared and restarted.");
  } catch (err) {
    await ctx.reply(`❌ Clear failed: ${err instanceof Error ? err.message : String(err)}`);
  }
}

export async function handleMenuButton(ctx: CommandContext): Promise<void> {
  await handleMenu(ctx);
}
