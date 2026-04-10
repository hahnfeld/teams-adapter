// Re-export activity tracking types from plugin-sdk
export { ToolStateMap, ThoughtBuffer, DisplaySpecBuilder } from "@openacp/plugin-sdk";

export type {
  OutputMode,
  ToolDisplaySpec,
  ToolCardSnapshot,
  PlanEntry,
  ToolCallMeta,
  ViewerLinks,
} from "@openacp/plugin-sdk";

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
