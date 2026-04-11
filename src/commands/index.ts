import type { TurnContext } from "@microsoft/agents-hosting";
import { log } from "@openacp/plugin-sdk";
import type { TeamsAdapter } from "../adapter.js";

import { handleNew, handleNewChat } from "./new-session.js";
import { handleCancel, handleStatus, handleSessions, handleHandoff } from "./session.js";
import { handleBypass, handleTTS, handleRestart, handleRespawn, handleUpdate, handleOutputMode } from "./admin.js";
import { handleMenu, handleHelp, handleClear } from "./menu.js";
import { handleAgents, handleInstall } from "./agents.js";
import { handleDoctor } from "./doctor.js";
import { handleIntegrate } from "./integrate.js";
import { handleSettings } from "./settings.js";

export interface CommandContext {
  context: TurnContext;
  adapter: TeamsAdapter;
  userId: string;
  sessionId: string | null;
  reply: (text: string) => Promise<void>;
}

const COMMAND_PREFIX = "/";

export const SLASH_COMMANDS = [
  { command: "new", description: "Create a new agent session" },
  { command: "newchat", description: "New chat, same agent & workspace" },
  { command: "cancel", description: "Cancel the current session" },
  { command: "status", description: "Show session or global status" },
  { command: "sessions", description: "List all sessions" },
  { command: "agents", description: "List available agents" },
  { command: "install", description: "Install an agent by name" },
  { command: "menu", description: "Show the action menu" },
  { command: "help", description: "Show help" },
  { command: "integrate", description: "Manage agent integrations" },
  { command: "handoff", description: "Generate terminal resume command" },
  { command: "restart", description: "Restart OpenACP" },
  { command: "respawn", description: "Restart the assistant session" },
  { command: "update", description: "Update to latest version" },
  { command: "doctor", description: "Run system diagnostics" },
  { command: "tts", description: "Toggle Text to Speech" },
  { command: "outputmode", description: "Set output detail level (low/medium/high)" },
  { command: "bypass", description: "Auto-approve permissions" },
  { command: "mode", description: "Switch session mode" },
  { command: "model", description: "Switch AI model" },
  { command: "thought", description: "Adjust thinking level" },
  { command: "settings", description: "Show configuration settings" },
  { command: "clear", description: "Reset the assistant session" },
];

export async function handleCommand(
  context: TurnContext,
  adapter: TeamsAdapter,
  userId: string,
  sessionId: string | null,
): Promise<boolean> {
  const text = (context.activity?.text ?? "").replace(/<at[^>]*>.*?<\/at>/gi, "").trim();
  if (!text.startsWith(COMMAND_PREFIX)) return false;

  const parts = text.slice(1).split(/\s+/);
  const commandName = parts[0].toLowerCase();
  const args = parts.slice(1);

  const ctx: CommandContext = {
    context,
    adapter,
    userId,
    sessionId,
    reply: async (content: string) => {
      if (typeof (context as any).send === "function") {
        await (context as any).send({ type: "message", text: content.replace(/(?<!\n)\n(?!\n)/g, "\n\n"), textFormat: "markdown" });
      } else {
        await (context.sendActivity as Function)({ text: content });
      }
    },
  };

  try {
    switch (commandName) {
      case "new":
        await handleNew(ctx, args);
        break;
      case "newchat":
        await handleNewChat(ctx);
        break;
      case "cancel":
        await handleCancel(ctx);
        break;
      case "status":
        await handleStatus(ctx);
        break;
      case "sessions":
        await handleSessions(ctx);
        break;
      case "agents":
        await handleAgents(ctx);
        break;
      case "install":
        await handleInstall(ctx, args[0]);
        break;
      case "menu":
        await handleMenu(ctx);
        break;
      case "help":
        await handleHelp(ctx);
        break;
      case "bypass":
        await handleBypass(ctx);
        break;
      case "restart":
        await handleRestart(ctx);
        break;
      case "respawn":
        await handleRespawn(ctx);
        break;
      case "update":
        await handleUpdate(ctx);
        break;
      case "integrate":
        await handleIntegrate(ctx);
        break;
      case "settings":
        await handleSettings(ctx);
        break;
      case "doctor":
        await handleDoctor(ctx);
        break;
      case "handoff":
        await handleHandoff(ctx);
        break;
      case "clear":
        await handleClear(ctx);
        break;
      case "tts":
        await handleTTS(ctx, args[0]);
        break;
      case "outputmode":
        await handleOutputMode(ctx, args[0], args[1]);
        break;
      case "verbosity":
        await handleOutputMode(ctx, args[0], args[1]);
        break;
      case "mode":
        await handleModeSwitch(ctx, args[0]);
        break;
      case "model":
        await handleModelSwitch(ctx, args[0]);
        break;
      case "thought":
        await handleThoughtLevel(ctx, args[0]);
        break;
      default:
        return false; // Not a locally handled command — let registry try
    }
    return true;
  } catch (err) {
    log.error({ err, commandName }, "[teams-router] Command handler failed");
    const errMsg = `❌ Command failed: ${err instanceof Error ? err.message : String(err)}`;
    try {
      await ctx.reply(errMsg);
    } catch { /* ignore */ }
    return true;
  }
}

export async function setupCardActionCallbacks(
  context: TurnContext,
  adapter: TeamsAdapter,
): Promise<void> {
  const value = context.activity.value as Record<string, unknown> | undefined;
  if (!value) return;

  // Support both Action.Execute (value.action.verb) and Action.Submit (value.verb) formats
  const action = value.action as { verb?: string; data?: Record<string, unknown> } | undefined;
  const verb = (action?.verb ?? value.verb) as string | undefined;
  const data = (action?.data ?? value) as Record<string, unknown>;

  if (!verb) return;

  const ctx: CommandContext = {
    context,
    adapter,
    userId: context.activity.from?.id ?? "unknown",
    sessionId: (data.sessionId as string | undefined) ?? null,
    reply: async (content: string) => {
      if (typeof (context as any).send === "function") {
        await (context as any).send({ type: "message", text: content.replace(/(?<!\n)\n(?!\n)/g, "\n\n"), textFormat: "markdown" });
      } else {
        await (context.sendActivity as Function)({ text: content });
      }
    },
  };

  // Route command verbs from Adaptive Card buttons
  if (verb.startsWith("cmd:")) {
    const command = verb.slice(4);
    try {
      switch (command) {
        case "menu":
          await handleMenu(ctx);
          break;
        case "help":
          await handleHelp(ctx);
          break;
        case "new":
          await handleNew(ctx, []);
          break;
        case "cancel":
          await handleCancel(ctx);
          break;
        case "status":
          await handleStatus(ctx);
          break;
        case "sessions":
          await handleSessions(ctx);
          break;
        case "agents":
          await handleAgents(ctx);
          break;
        case "doctor":
          await handleDoctor(ctx);
          break;
        case "noop":
          break;
        default:
          await ctx.reply(`Unknown action: ${command}`);
      }
    } catch (err) {
      log.error({ err, command }, "[teams-router] Card action command failed");
      await ctx.reply(`❌ Command failed`);
    }
  }
}

// ─── Session config commands ──────────────────────────────────────────────

/**
 * Delegate a session-scoped command to the adapter's handleCommand, which
 * dispatches through the CommandRegistry. Avoids duplicating registry
 * lookup logic that the adapter already handles.
 */
async function delegateToRegistry(ctx: CommandContext, commandText: string, fallbackMsg: string): Promise<void> {
  if (!ctx.sessionId) {
    await ctx.reply("❌ No active session.");
    return;
  }
  try {
    await ctx.adapter.handleCommand(commandText, ctx.context, ctx.sessionId, ctx.userId);
  } catch {
    await ctx.reply(fallbackMsg);
  }
}

async function handleModeSwitch(ctx: CommandContext, mode?: string): Promise<void> {
  if (!mode) {
    await ctx.reply("Usage: `/mode <mode-name>`\n\nExample: `/mode plan`, `/mode code`");
    return;
  }
  await delegateToRegistry(ctx, `/mode ${mode}`, `🔄 Mode set to **${mode}** (may require core command support)`);
}

async function handleModelSwitch(ctx: CommandContext, model?: string): Promise<void> {
  if (!model) {
    await ctx.reply("Usage: `/model <model-name>`\n\nExample: `/model claude-sonnet`, `/model gpt-4o`");
    return;
  }
  await delegateToRegistry(ctx, `/model ${model}`, `🤖 Model set to **${model}** (may require core command support)`);
}

async function handleThoughtLevel(ctx: CommandContext, level?: string): Promise<void> {
  if (!level) {
    await ctx.reply("Usage: `/thought <level>`\n\nExample: `/thought high`, `/thought low`, `/thought off`");
    return;
  }
  await delegateToRegistry(ctx, `/thought ${level}`, `🧠 Thinking level set to **${level}** (may require core command support)`);
}