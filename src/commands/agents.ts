import { sendInfoCard } from "./index.js";
import type { CommandContext } from "./index.js";
import { log } from "@openacp/plugin-sdk";

/**
 * Handle /agents — list installed and available agents.
 */
export async function handleAgents(ctx: CommandContext): Promise<void> {
  const catalog = ctx.adapter.core.agentCatalog;
  if (!catalog) {
    await sendInfoCard(ctx, "❌", "Error", "Agent catalog not available.");
    return;
  }

  const items = catalog.getAvailable();
  const installed = items.filter((i) => i.installed);
  const available = items.filter((i) => !i.installed);

  const lines: string[] = [];

  if (installed.length > 0) {
    lines.push("Installed:");
    for (const item of installed) {
      const desc = item.description ? ` — ${truncate(item.description, 50)}` : "";
      lines.push(`✓ ${item.name}${desc}`);
    }
  }

  if (available.length > 0) {
    if (installed.length > 0) lines.push("");
    lines.push("Available:");
    for (const item of available.slice(0, 10)) {
      if (item.available) {
        lines.push(`◻ ${item.name}`);
      } else {
        const deps = item.missingDeps?.join(", ") ?? "requirements not met";
        lines.push(`⚠ ${item.name} (needs: ${deps})`);
      }
    }
    if (available.length > 10) {
      lines.push(`...and ${available.length - 10} more`);
    }
  } else if (installed.length > 0) {
    lines.push("All agents installed!");
  }

  await sendInfoCard(ctx, "🤖", "Agents", lines.join("\n"));
}

/**
 * Handle /install <agent> — install an agent by name.
 */
export async function handleInstall(ctx: CommandContext, name?: string): Promise<void> {
  if (!name) {
    await sendInfoCard(ctx, "📦", "Install", "Usage: /install <agent-name>");
    return;
  }

  const catalog = ctx.adapter.core.agentCatalog;
  if (!catalog) {
    await sendInfoCard(ctx, "❌", "Error", "Agent catalog not available.");
    return;
  }

  await sendInfoCard(ctx, "⏳", "Installing", name);

  try {
    const result = await catalog.install(name, {
      onStart: () => {},
      onStep: async () => {},
      onDownloadProgress: async () => {},
      onSuccess: async () => {},
      onError: async () => {},
    });

    if (result.ok) {
      const steps = result.setupSteps?.length
        ? `\nSetup: ${result.setupSteps.join(", ")}`
        : "";
      await sendInfoCard(ctx, "✅", "Installed", `${name}${steps}`);
    } else {
      await sendInfoCard(ctx, "❌", "Install failed", result.error ?? "Unknown error");
    }
  } catch (err) {
    log.error({ err, name }, "[agents] Install failed");
    await sendInfoCard(ctx, "❌", "Install failed", err instanceof Error ? err.message : String(err));
  }
}

function truncate(text: string, maxLen: number): string {
  return text.length <= maxLen ? text : text.slice(0, maxLen - 1) + "…";
}
