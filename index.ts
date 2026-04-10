import type { OpenACPPlugin } from "@openacp/plugin-sdk";
import { TeamsAdapter } from "./src/adapter.js";
import type { TeamsChannelConfig } from "./src/types.js";

export default function createTeamsPlugin(): OpenACPPlugin {
  let adapter: TeamsAdapter | null = null;

  return {
    name: "@openacp/teams",
    version: "1.0.0",
    essential: false,

    async setup(ctx) {
      const config = ctx.configManager.get().channels?.teams as TeamsChannelConfig | undefined;
      if (!config?.enabled) {
        return;
      }
      const { TeamsAdapter } = await import("./src/adapter.js");
      adapter = new TeamsAdapter(ctx.core, config);
      ctx.registerService("adapter:teams", adapter);
    },

    async teardown() {
      if (adapter) {
        await adapter.stop();
        adapter = null;
      }
    },
  };
}