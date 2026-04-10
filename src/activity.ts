// Re-export activity tracking types from plugin-sdk
// Note: Base ActivityTracker from plugin-sdk has a different API than the Discord version
// Teams adapter uses inline activity handling in adapter.ts

export { ToolStateMap, ThoughtBuffer, DisplaySpecBuilder } from "@openacp/plugin-sdk";

export type {
  OutputMode,
  ToolDisplaySpec,
  ToolCardSnapshot,
  PlanEntry,
  ToolCallMeta,
  ViewerLinks,
} from "@openacp/plugin-sdk";

export interface SendQueue {
  enqueue<T>(fn: () => Promise<T>, opts?: { type?: string }): Promise<T | undefined>;
}

export interface MessageRef {
  activityId?: string;
  conversationId?: string;
}

export interface ToolEntry {
  id: string;
  name: string;
  kind: string;
  rawInput: unknown;
  content: string | null;
  status: string;
  viewerLinks?: { file?: string; diff?: string };
  diffStats?: { added: number; removed: number };
  displaySummary?: string;
  displayTitle?: string;
  displayKind?: string;
  isNoise: boolean;
}