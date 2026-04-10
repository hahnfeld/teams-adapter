import type { AdaptiveCards } from "@microsoft/teams.cards";

// TODO: Import these from @openacp/plugin-sdk once the SDK is updated with
// the new OutputMode / ToolDisplaySpec / ToolCardSnapshot exports.
// For now, defined locally to unblock development.

export type OutputMode = "low" | "medium" | "high";

export interface ToolDisplaySpec {
  id: string;
  kind: string;
  icon: string;
  title: string;
  description: string | null;
  command: string | null;
  inputContent: string | null;
  outputSummary: string | null;
  outputContent: string | null;
  diffStats: { added: number; removed: number } | null;
  viewerLinks?: { file?: string; diff?: string };
  outputViewerLink?: string;
  outputFallbackContent?: string;
  status: string;
  isNoise: boolean;
  isHidden: boolean;
}

export interface PlanEntry {
  content: string;
  status: string;
  priority: string;
}

export interface ToolCardSnapshot {
  specs: ToolDisplaySpec[];
  planEntries?: PlanEntry[];
  usage?: { tokensUsed?: number; contextSize?: number; cost?: number };
  totalVisible: number;
  completedVisible: number;
  allComplete: boolean;
}

export interface ToolCardResult {
  body: AdaptiveCards.CardElement[];
  actions: AdaptiveCards.ActionElement[];
}

// ─── Constants (local copies; TODO: import from @openacp/plugin-sdk/formatting) ─

const STATUS_ICONS: Record<string, string> = {
  pending: "⏳",
  in_progress: "🔄",
  completed: "✅",
  failed: "❌",
  cancelled: "🚫",
  running: "🔄",
  done: "✅",
  error: "❌",
};

const KIND_ICONS: Record<string, string> = {
  read: "📖",
  edit: "✏️",
  write: "✏️",
  delete: "🗑️",
  execute: "▶️",
  command: "▶️",
  bash: "▶️",
  terminal: "▶️",
  search: "🔍",
  web: "🌐",
  fetch: "🌐",
  agent: "🧠",
  think: "🧠",
  install: "📦",
  move: "📦",
  other: "🛠️",
};

const KIND_LABELS: Record<string, string> = {
  read: "Read",
  edit: "Edit",
  write: "Write",
  delete: "Delete",
  execute: "Run",
  bash: "Bash",
  command: "Run",
  terminal: "Terminal",
  search: "Search",
  web: "Web",
  fetch: "Fetch",
  agent: "Agent",
  think: "Agent",
  install: "Install",
  move: "Move",
};

// ─── Helpers ────────────────────────────────────────────────────────────────

function progressBar(ratio: number, length = 10): string {
  const filled = Math.round(Math.min(1, Math.max(0, ratio)) * length);
  return "▓".repeat(filled) + "░".repeat(length - filled);
}

function formatTokens(n: number): string {
  if (n >= 1000) return `${Math.round(n / 1000)}k`;
  return String(n);
}

const TRUNCATION_SUFFIX = "… (truncated)";

function truncateContent(text: string, max: number): string {
  if (text.length <= max) return text;
  return text.slice(0, max - TRUNCATION_SUFFIX.length) + TRUNCATION_SUFFIX;
}

const INLINE_OUTPUT_MAX = 800;

// ─── renderSpecSection ──────────────────────────────────────────────────────

export function renderSpecSection(spec: ToolDisplaySpec, mode: OutputMode): string {
  const statusIcon = STATUS_ICONS[spec.status] ?? "🔧";
  const kindIcon = spec.icon || KIND_ICONS[spec.kind] || "🔧";
  const kindLabel = KIND_LABELS[spec.kind] || "";

  if (mode === "low") {
    return `${statusIcon} ${kindIcon} ${kindLabel || spec.title}`;
  }

  const lines: string[] = [];

  const leadIcon = spec.isNoise && mode === "high" ? "👁️" : statusIcon;
  const titleLine = `${leadIcon} ${kindIcon} **${spec.title}**`;
  lines.push(titleLine);

  if (spec.description) {
    lines.push(` ╰ ${spec.description}`);
  }

  const diffParts: string[] = [];
  if (spec.diffStats) {
    const { added, removed } = spec.diffStats;
    if (added > 0 && removed > 0) diffParts.push(`+${added}/-${removed} lines`);
    else if (added > 0) diffParts.push(`+${added} lines`);
    else if (removed > 0) diffParts.push(`-${removed} lines`);
  }
  if (spec.viewerLinks?.diff) {
    diffParts.push(`[View Diff](${spec.viewerLinks.diff})`);
  }
  if (spec.viewerLinks?.file) {
    diffParts.push(`[View File](${spec.viewerLinks.file})`);
  }
  if (diffParts.length > 0) {
    lines.push(` ╰ ${diffParts.join(" · ")}`);
  }

  if (spec.outputSummary && !spec.outputContent) {
    lines.push(` ╰ ${spec.outputSummary}`);
  }

  if (mode === "high") {
    if (spec.outputContent || spec.outputFallbackContent) {
      const raw = spec.outputContent ?? spec.outputFallbackContent!;
      const truncated = truncateContent(raw, INLINE_OUTPUT_MAX);
      lines.push(`\`\`\`\n${truncated}\n\`\`\``);
    }
    if (spec.outputViewerLink) {
      lines.push(`[View output](${spec.outputViewerLink})`);
    }
  }

  return lines.join("\n");
}

// ─── renderToolCard ──────────────────────────────────────────────────────────

export function renderToolCard(
  snapshot: ToolCardSnapshot,
  mode: OutputMode,
  sessionId?: string,
  thoughtViewerLink?: string,
): ToolCardResult {
  const { specs, totalVisible, completedVisible, allComplete } = snapshot;

  const visible = specs.filter((s) => !s.isHidden);

  const hasError = visible.some((s) => s.status === "error" || s.status === "failed");
  const cardColor = hasError ? "#e74c3c" : allComplete ? "#2ecc71" : "#3498db";

  const authorName = allComplete
    ? `✅ Done ${completedVisible}/${totalVisible}`
    : `🔄 Working... ${completedVisible} of ${totalVisible}`;

  const sections = visible.map((s) => renderSpecSection(s, mode));
  let description: string;

  if (mode === "low") {
    const lines: string[] = [];
    for (let i = 0; i < sections.length; i += 3) {
      lines.push(sections.slice(i, i + 3).join(" · "));
    }
    description = lines.join("\n");
  } else {
    description = sections.join("\n\n");
  }

  if (snapshot.planEntries && snapshot.planEntries.length > 0) {
    const entries = snapshot.planEntries;
    if (mode === "high") {
      const planLines = entries.map(
        (e, i) => `${STATUS_ICONS[e.status] ?? "⬜"} ${i + 1}. ${e.content}`,
      );
      description += "\n\n📋 **Plan:**\n" + planLines.join("\n");
    }
  }

  if (mode === "high" && thoughtViewerLink) {
    description += `\n\n💭 [View Thinking](${thoughtViewerLink})`;
  }

  if (!description) {
    return { body: [], actions: [] };
  }

  const body: AdaptiveCards.CardElement[] = [
    {
      type: "TextBlock",
      text: authorName,
      weight: "Bolder",
      size: "Medium",
      color: hasError ? "Attention" : allComplete ? "Good" : "Accent",
    },
    {
      type: "TextBlock",
      text: description,
      wrap: true,
      spacing: "Medium",
    },
  ];

  if (mode === "medium" && snapshot.planEntries?.length) {
    const entries = snapshot.planEntries;
    const currentIdx = entries.findIndex((e) => e.status === "in_progress");
    const stepNum = currentIdx >= 0 ? currentIdx + 1 : entries.filter((e) => e.status === "completed").length + 1;
    const currentLabel = currentIdx >= 0 ? entries[currentIdx].content : entries[Math.min(stepNum - 1, entries.length - 1)]?.content ?? "";
    body.push({
      type: "TextBlock",
      text: `📋 Step ${stepNum}/${entries.length} — ${currentLabel}`,
      size: "Small",
      isSubtle: true,
    });
  }

  const actions: AdaptiveCards.ActionElement[] = [];
  if (!allComplete && sessionId) {
    actions.push(
      { type: "Action.Execute", title: "🔇 Low", data: { action: `om:${sessionId}:low` } },
      { type: "Action.Execute", title: "📊 Medium", data: { action: `om:${sessionId}:medium` } },
      { type: "Action.Execute", title: "🔍 High", data: { action: `om:${sessionId}:high` } },
      { type: "Action.Execute", title: "❌ Cancel", data: { action: `cancel:${sessionId}` } },
    );
  }

  return { body, actions };
}

// ─── renderUsageCard ────────────────────────────────────────────────────────

interface UsageData {
  tokensUsed?: number;
  contextSize?: number;
  cost?: number;
  duration?: number;
}

export function renderUsageCard(usage: UsageData, mode: OutputMode): { body: AdaptiveCards.CardElement[] } {
  const { tokensUsed, contextSize, cost, duration } = usage;

  if (tokensUsed == null) {
    return { body: [{ type: "TextBlock", text: "📊 Usage data unavailable", wrap: true }] };
  }

  const durationStr = duration != null ? `${duration}s` : null;

  if (mode === "low") {
    const parts = [`📊 ${formatTokens(tokensUsed)} tokens`];
    if (durationStr) parts.push(durationStr);
    return { body: [{ type: "TextBlock", text: parts.join(" · "), wrap: true }] };
  }

  if (mode === "medium") {
    const line1Parts = [`📊 ${formatTokens(tokensUsed)} tokens`];
    if (cost != null) line1Parts.push(`$${cost.toFixed(2)}`);
    const lines = [line1Parts.join(" · ")];
    if (durationStr) lines.push(`⏱️ ${durationStr}`);
    return { body: [{ type: "TextBlock", text: lines.join("\n"), wrap: true }] };
  }

  if (contextSize == null) {
    return { body: [{ type: "TextBlock", text: `📊 ${formatTokens(tokensUsed)} tokens`, wrap: true }] };
  }

  const ratio = tokensUsed / contextSize;
  const pct = Math.round(ratio * 100);
  const bar = progressBar(ratio);
  const emoji = pct >= 85 ? "⚠️" : "📊";

  const lines = [`${emoji} ${formatTokens(tokensUsed)} / ${formatTokens(contextSize)} tokens`, `${bar} ${pct}%`];
  if (cost != null) lines.push(`💰 $${cost.toFixed(2)}`);
  if (durationStr) lines.push(`⏱️ ${durationStr}`);

  return { body: [{ type: "TextBlock", text: lines.join("\n"), wrap: true }] };
}

// ─── renderPermissionCard ────────────────────────────────────────────────────

interface PermissionRequest {
  toolName: string;
  command?: string;
  description?: string;
}

export function renderPermissionCard(
  request: PermissionRequest,
  sessionId: string,
  callbackKey: string,
): { body: AdaptiveCards.CardElement[]; actions: AdaptiveCards.ActionElement[] } {
  const prefix = `p:${sessionId}:${callbackKey}`;

  const descParts: string[] = [];
  descParts.push(`**Tool:** ${request.toolName}`);
  if (request.command) {
    descParts.push(`**Command:** \`${request.command}\``);
  }
  if (request.description) {
    descParts.push(request.description);
  }

  const body: AdaptiveCards.CardElement[] = [
    {
      type: "TextBlock",
      text: "🔐 Permission Request",
      weight: "Bolder",
      color: "Warning",
    },
    {
      type: "TextBlock",
      text: descParts.join("\n"),
      wrap: true,
      spacing: "Medium",
    },
  ];

  const actions: AdaptiveCards.ActionElement[] = [
    { type: "Action.Execute", title: "Allow", data: { verb: "allow", sessionId, callbackKey: prefix } },
    { type: "Action.Execute", title: "Deny", data: { verb: "deny", sessionId, callbackKey: prefix } },
    { type: "Action.Execute", title: "Always Allow", data: { verb: "always", sessionId, callbackKey: prefix } },
  ];

  return { body, actions };
}

// ─── Format helpers (legacy) ─────────────────────────────────────────────────

function extractContentTextLegacy(content: unknown): string {
  if (content == null) return "";
  if (typeof content === "string") return content;
  if (Array.isArray(content)) {
    return content
      .map((c) => {
        if (typeof c === "string") return c;
        if (c && typeof c === "object" && "text" in c) return String(c.text);
        return "";
      })
      .filter(Boolean)
      .join("\n");
  }
  if (typeof content === "object" && content !== null && "text" in content) {
    return String((content as { text: unknown }).text);
  }
  return "";
}

function stripCodeFencesLegacy(text: string): string {
  return text.replace(/^```[\w]*\n?/gm, "").replace(/\n?```$/gm, "");
}

function formatViewerLinksLegacy(links?: { file?: string; diff?: string }, filePath?: string): string {
  if (!links) return "";
  const fileName = filePath ? filePath.split("/").pop() || filePath : "";
  let text = "\n";
  if (links.file) text += `\n[View ${fileName || "file"}](${links.file})`;
  if (links.diff) text += `\n[View diff${fileName ? ` — ${fileName}` : ""}](${links.diff})`;
  return text;
}

function formatHighDetailsLegacy(
  rawInput: unknown,
  content: unknown,
  maxLen: number,
): string {
  let text = "";
  if (rawInput) {
    const inputStr = typeof rawInput === "string" ? rawInput : JSON.stringify(rawInput, null 2);
    if (inputStr && inputStr !== "{}") {
      text += `\n**Input:**\n\`\`\`\n${truncateContent(inputStr, maxLen)}\n\`\`\``;
    }
  }
  const details = stripCodeFencesLegacy(extractContentTextLegacy(content));
  if (details) {
    text += `\n**Output:**\n\`\`\`\n${truncateContent(details, maxLen)}\n\`\`\``;
  }
  return text;
}

type LegacyVerbosity = "low" | "medium" | "high";

interface LegacyToolCallMeta {
  id: string;
  name: string;
  kind?: string;
  status?: string;
  content?: unknown;
  rawInput?: unknown;
  viewerLinks?: { file?: string; diff?: string };
  viewerFilePath?: string;
  displaySummary?: string;
  displayTitle?: string;
  displayKind?: string;
}

function legacyResolveToolIcon(tool: LegacyToolCallMeta): string {
  if (tool.status && STATUS_ICONS[tool.status]) return STATUS_ICONS[tool.status];
  if (tool.displayKind && KIND_ICONS[tool.displayKind]) return KIND_ICONS[tool.displayKind];
  if (tool.kind && KIND_ICONS[tool.kind]) return KIND_ICONS[tool.kind];
  return "🔧";
}

function legacyFormatTitle(name: string, _rawInput: unknown, displayTitle?: string): string {
  if (displayTitle) return displayTitle;
  return name;
}

function legacyFormatSummary(name: string, rawInput: unknown, displaySummary?: string): string {
  if (displaySummary) return displaySummary;
  if (rawInput && typeof rawInput === "object") {
    const input = rawInput as Record<string, unknown>;
    if (input.pattern) return `${KIND_ICONS[name.toLowerCase()] || "🔍"} ${name} "${input.pattern}"`;
    if (input.file_path) return `${name} ${input.file_path}`;
  }
  return name;
}

/** @deprecated Use renderToolCard instead */
export function formatToolCall(
  tool: LegacyToolCallMeta,
  verbosity: LegacyVerbosity = "medium",
): string {
  const si = legacyResolveToolIcon(tool);
  const name = tool.name || "Tool";
  const label = verbosity === "low"
    ? legacyFormatTitle(name, tool.rawInput, tool.displayTitle)
    : legacyFormatSummary(name, tool.rawInput, tool.displaySummary);
  let text = `${si} **${label}**`;
  text += formatViewerLinksLegacy(tool.viewerLinks, tool.viewerFilePath);
  if (verbosity === "high") {
    text += formatHighDetailsLegacy(tool.rawInput, tool.content, 500);
  }
  return text;
}

/** @deprecated Use renderToolCard instead */
export function formatToolUpdate(update: LegacyToolCallMeta, verbosity: LegacyVerbosity = "medium"): string {
  return formatToolCall(update, verbosity);
}

/** @deprecated Use renderToolCard plan rendering instead */
export function formatPlan(entries: PlanEntry[], verbosity: LegacyVerbosity = "medium"): string {
  if (verbosity === "medium") {
    const done = entries.filter((e) => e.status === "completed").length;
    return `📋 **Plan:** ${done}/${entries.length} steps completed`;
  }
  const statusIconMap: Record<string, string> = {
    pending: "⏳",
    in_progress: "🔄",
    completed: "✅",
  };
  const lines = entries.map((e, i) => `${statusIconMap[e.status] || "⬜"} ${i + 1}. ${e.content}`);
  return `**Plan:**\n${lines.join("\n")}`;
}

/** @deprecated Use renderUsageCard instead */
export function formatUsage(
  usage: { tokensUsed?: number; contextSize?: number; cost?: number },
  verbosity: LegacyVerbosity = "medium",
): string {
  const { tokensUsed, contextSize, cost } = usage;
  if (tokensUsed == null) return "📊 Usage data unavailable";
  if (verbosity === "medium") {
    const costStr = cost != null ? ` · $${cost.toFixed(2)}` : "";
    return `📊 ${formatTokens(tokensUsed)} tokens${costStr}`;
  }
  if (contextSize == null) return `📊 ${formatTokens(tokensUsed)} tokens`;
  const ratio = tokensUsed / contextSize;
  const pct = Math.round(ratio * 100);
  const bar = progressBar(ratio);
  const emoji = pct >= 85 ? "⚠️" : "📊";
  let text = `${emoji} ${formatTokens(tokensUsed)} / ${formatTokens(contextSize)} tokens\n${bar} ${pct}%`;
  if (cost != null) text += `\n💰 $${cost.toFixed(2)}`;
  return text;
}

function splitMessageImpl(text: string, maxLength: number): string[] {
  if (text.length <= maxLength) return [text];
  const paragraphs = text.split("\n\n");
  const chunks: string[] = [];
  let current = "";
  for (const para of paragraphs) {
    const candidate = current ? `${current}\n\n${para}` : para;
    if (candidate.length > maxLength && current) {
      chunks.push(current);
      current = para.length > maxLength ? para.slice(0, maxLength) : para;
    } else if (candidate.length > maxLength) {
      const lines = para.split("\n");
      for (const line of lines) {
        const lineCandidate = current ? `${current}\n${line}` : line;
        if (lineCandidate.length > maxLength && current) {
          chunks.push(current);
          current = line;
        } else {
          current = lineCandidate;
        }
      }
    } else {
      current = candidate;
    }
  }
  if (current) chunks.push(current);
  return chunks;
}

/** @deprecated Use splitMessageImpl instead */
export function splitMessage(text: string, maxLength = 1800): string[] {
  return splitMessageImpl(text, maxLength);
}