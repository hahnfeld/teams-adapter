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
): Promise<void> {
  const text = context.activity.text ?? "";
  if (!text.startsWith(COMMAND_PREFIX)) return;

  const parts = text.slice(1).split(/\s+/);
  const commandName = parts[0].toLowerCase();
  const args = parts.slice(1);

  const ctx: CommandContext = {
    context,
    adapter,
    userId,
    sessionId,
    reply: async (content: string) => {
      await context.sendActivity(content as any);
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
        await ctx.reply("Mode switching not yet implemented");
        break;
      case "model":
        await ctx.reply("Model switching not yet implemented");
        break;
      case "thought":
        await ctx.reply("Thought level not yet implemented");
        break;
      default:
        await ctx.reply(`Unknown command: /${commandName}`);
    }
  } catch (err) {
    log.error({ err, commandName }, "[teams-router] Command handler failed");
    const errMsg = `❌ Command failed: ${err instanceof Error ? err.message : String(err)}`;
    try {
      await context.sendActivity(errMsg as any);
    } catch { /* ignore */ }
  }
}

export async function setupCardActionCallbacks(
  context: TurnContext,
  adapter: TeamsAdapter,
): Promise<void> {
  const value = context.activity.value as { action?: { verb?: string; data?: Record<string, unknown> } } | undefined;
  const action = value?.action;
  if (!action) return;

  const verb = action.verb as string;
  const data = (action.data ?? {}) as Record<string, unknown>;

  const ctx: CommandContext = {
    context,
    adapter,
    userId: context.activity.from?.id ?? "unknown",
    sessionId: (data.sessionId as string | undefined) ?? null,
    reply: async (content: string) => {
      await context.sendActivity(content as any);
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
        case "noop":
          // No-op for cancel dialog negative button
          break;
        default:
          await ctx.reply(`Command not yet implemented: ${command}`);
      }
    } catch (err) {
      log.error({ err, command }, "[teams-router] Card action command failed");
      await ctx.reply(`❌ Command failed`);
    }
  }
}