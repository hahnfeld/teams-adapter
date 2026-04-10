import type { OpenACPPlugin } from "@openacp/plugin-sdk";
import { log } from "@openacp/plugin-sdk";
import { TeamsAdapter } from "./adapter.js";
import type { TeamsChannelConfig } from "./types.js";

/**
 * Factory for the Teams adapter plugin.
 *
 * Follows the Telegram plugin pattern:
 * - Uses `ctx.pluginConfig` for adapter-specific config (not ctx.configManager)
 * - Registers via `ctx.registerService("adapter:teams", ...)` (not ctx.registerAdapter)
 * - Does NOT call adapter.start() — core's lifecycle manager handles that
 */
export default function createTeamsPlugin(): OpenACPPlugin {
  let adapter: TeamsAdapter | null = null;

  return {
    name: "@openacp/teams",
    version: "1.0.0",
    essential: false,
    permissions: ["services:register", "kernel:access", "events:read"],

    async setup(ctx) {
      // Read Teams config from pluginConfig (matches Telegram pattern)
      const config = ctx.pluginConfig as Record<string, unknown>;
      if (!config.enabled || !config.botAppId) {
        ctx.log.info("Teams adapter disabled (missing enabled or botAppId)");
        return;
      }

      const core = ctx.core as import("@openacp/plugin-sdk").OpenACPCore;

      adapter = new TeamsAdapter(core, config as unknown as TeamsChannelConfig);
      // Register as a named service — core discovers adapters via "adapter:*" keys
      ctx.registerService("adapter:teams", adapter);
      ctx.log.info("Teams adapter registered");
    },

    async teardown() {
      if (adapter) {
        await adapter.stop();
        adapter = null;
      }
    },
  };
}
