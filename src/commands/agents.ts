import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";

/**
 * Handle /agents — list installed and available agents.
 * Mirrors Telegram's handleAgents with install buttons.
 */
export async function handleAgents(ctx: CommandContext): Promise<void> {
  const catalog = ctx.adapter.core.agentCatalog;
  if (!catalog) {
    await ctx.reply("❌ Agent catalog not available.");
    return;
  }

  const items = catalog.getAvailable();
  const installed = items.filter((i) => i.installed);
  const available = items.filter((i) => !i.installed);

  let text = "**🤖 Agents**\n\n";

  if (installed.length > 0) {
    text += "**Installed:**\n\n";
    for (const item of installed) {
      text += `✅ **${item.name}**`;
      if (item.description) text += ` — ${truncate(item.description, 50)}`;
      text += "\n\n";
    }
  }

  if (available.length > 0) {
    if (installed.length > 0) text += "---\n\n";
    text += "**Available to install:**\n\n";
    for (const item of available.slice(0, 10)) {
      if (item.available) {
        text += `⬇️ **${item.name}**`;
      } else {
        const deps = item.missingDeps?.join(", ") ?? "requirements not met";
        text += `⚠️ **${item.name}** (needs: ${deps})`;
      }
      if (item.description) text += `\n\n    ${truncate(item.description, 60)}`;
      text += "\n\n";
    }
    if (available.length > 10) {
      text += `...and ${available.length - 10} more. Use \`/install <name>\` to install.`;
    }
  } else if (installed.length > 0) {
    text += "All agents are already installed!";
  }

  await ctx.reply(text);
}

/**
 * Handle /install <agent> — install an agent by name.
 * Mirrors Telegram's installAgentWithProgress.
 */
export async function handleInstall(ctx: CommandContext, name?: string): Promise<void> {
  if (!name) {
    await ctx.reply("**Install an agent**\n\nUsage: `/install <agent-name>`\n\nUse `/agents` to browse available agents.");
    return;
  }

  const catalog = ctx.adapter.core.agentCatalog;
  if (!catalog) {
    await ctx.reply("❌ Agent catalog not available.");
    return;
  }

  await ctx.reply(`⏳ Installing **${name}**...`);

  try {
    const result = await catalog.install(name, {
      onStart: () => {},
      onStep: async () => {},
      onDownloadProgress: async () => {},
      onSuccess: async () => {},
      onError: async () => {},
    });

    if (result.ok) {
      let msg = `✅ **${name}** installed!`;
      if (result.setupSteps?.length) {
        msg += "\n\n**Setup steps:**\n\n";
        for (const step of result.setupSteps) {
          msg += `- ${step}\n\n`;
        }
      }
      await ctx.reply(msg);
    } else {
      await ctx.reply(`❌ Installation failed: ${result.error ?? "Unknown error"}`);
    }
  } catch (err) {
    log.error({ err, name }, "[agents] Install failed");
    await ctx.reply(`❌ Failed to install: ${err instanceof Error ? err.message : String(err)}`);
  }
}

export async function handleAgentButton(ctx: CommandContext): Promise<void> {
  await handleAgents(ctx);
}

function truncate(text: string, maxLen: number): string {
  return text.length <= maxLen ? text : text.slice(0, maxLen - 1) + "…";
}
