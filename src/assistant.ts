import type { OpenACPCore } from "@openacp/plugin-sdk";

export async function spawnAssistant(
  core: OpenACPCore,
  threadId: string,
): Promise<{ session: unknown; ready: Promise<unknown> }> {
  // TODO: Implement Teams assistant session spawning
  // Similar to Discord's spawnAssistant in assistant.ts
  throw new Error("Not yet implemented");
}

export function buildWelcomeMessage(core: OpenACPCore): string {
  return "👋 Welcome to OpenACP on Microsoft Teams!";
}