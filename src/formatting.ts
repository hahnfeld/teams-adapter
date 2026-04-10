// Formatting exports — helpers from plugin-sdk where available
export type { OutputMode, ToolDisplaySpec, ToolCardSnapshot, PlanEntry } from "@openacp/plugin-sdk";

// Status/Kind constants
export const STATUS_ICONS: Record<string, string> = {
  pending: "⏳", in_progress: "🔄", completed: "✅", failed: "❌",
  cancelled: "🚫", running: "🔄", done: "✅", error: "❌",
};

export const KIND_ICONS: Record<string, string> = {
  read: "📖", edit: "✏️", write: "✏️", delete: "🗑️", execute: "▶️",
  command: "▶️", bash: "▶️", terminal: "▶️", search: "🔍", web: "🌐",
  fetch: "🌐", agent: "🧠", think: "🧠", install: "📦", move: "📦", other: "🛠️",
};

export function progressBar(ratio: number, length = 10): string {
  const filled = Math.round(Math.min(1, Math.max(0, ratio)) * length);
  return "▓".repeat(filled) + "░".repeat(length - filled);
}

export function formatTokens(n: number): string {
  if (n >= 1000) return `${Math.round(n / 1000)}k`;
  return String(n);
}

export function truncateContent(text: string, max: number): string {
  const suffix = "… (truncated)";
  if (text.length <= max) return text;
  return text.slice(0, max - suffix.length) + suffix;
}

export function splitMessage(text: string, maxLength: number): string[] {
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

export function extractContentText(content: unknown): string {
  if (content == null) return "";
  if (typeof content === "string") return content;
  if (Array.isArray(content)) {
    return content.map((c) => {
      if (typeof c === "string") return c;
      if (c && typeof c === "object" && "text" in c) return String((c as { text: unknown }).text);
      return "";
    }).filter(Boolean).join("\n");
  }
  if (typeof content === "object" && content !== null && "text" in content) {
    return String((content as { text: unknown }).text);
  }
  return "";
}

export function stripCodeFences(text: string): string {
  return text.replace(/^```[\w]*\n?/gm, "").replace(/\n?```$/gm, "");
}

export function resolveToolIcon(kind: string, displayKind?: string, status?: string): string {
  if (status && STATUS_ICONS[status]) return STATUS_ICONS[status];
  if (displayKind && KIND_ICONS[displayKind]) return KIND_ICONS[displayKind];
  if (kind && KIND_ICONS[kind]) return KIND_ICONS[kind];
  return "🔧";
}

export function formatToolTitle(name: string, _rawInput?: unknown, displayTitle?: string): string {
  if (displayTitle) return displayTitle;
  return name;
}

export function formatToolSummary(name: string, rawInput?: unknown, displaySummary?: string): string {
  if (displaySummary) return displaySummary;
  if (rawInput && typeof rawInput === "object") {
    const input = rawInput as Record<string, unknown>;
    if (input.pattern) return `${name} "${input.pattern}"`;
    if (input.file_path) return `${name} ${input.file_path}`;
  }
  return name;
}

export function formatToolCall(
  tool: { id: string; name: string; kind?: string; status?: string; content?: unknown; rawInput?: unknown; viewerLinks?: { file?: string; diff?: string }; viewerFilePath?: string; displaySummary?: string; displayTitle?: string; displayKind?: string },
  verbosity: "low" | "medium" | "high" = "medium",
): string {
  const si = resolveToolIcon(tool.kind ?? "", tool.displayKind, tool.status);
  const name = tool.name || "Tool";
  const label = verbosity === "low"
    ? formatToolTitle(name, tool.rawInput, tool.displayTitle)
    : formatToolSummary(name, tool.rawInput, tool.displaySummary);
  let text = `${si} **${label}**`;
  if (tool.viewerLinks) {
    const fn = tool.viewerFilePath ? tool.viewerFilePath.split("/").pop() || tool.viewerFilePath : "";
    const fileLink = tool.viewerLinks.file;
    const diffLink = tool.viewerLinks.diff;
    if (fileLink) {
      text += "\n[View " + (fn || "file") + "](" + fileLink + ")";
    }
    if (diffLink) {
      text += "\n[View diff" + (fn ? " - " + fn : "") + "](" + diffLink + ")";
    }
  }
  if (verbosity === "high" && (tool.rawInput || tool.content)) {
    const maxLen = 500;
    let detail = "";
    if (tool.rawInput) {
      const inputStr = typeof tool.rawInput === "string" ? tool.rawInput : JSON.stringify(tool.rawInput, null, 2);
      if (inputStr && inputStr !== "{}") {
        detail += `\n**Input:**\n\`\`\`\n${truncateContent(inputStr, maxLen)}\n\`\`\``;
      }
    }
    const contentText = stripCodeFences(extractContentText(tool.content));
    if (contentText) {
      detail += `\n**Output:**\n\`\`\`\n${truncateContent(contentText, maxLen)}\n\`\`\``;
    }
    text += detail;
  }
  return text;
}

export function formatToolUpdate(
  tool: { id: string; name: string; kind?: string; status?: string; content?: unknown; rawInput?: unknown; viewerLinks?: { file?: string; diff?: string }; viewerFilePath?: string; displaySummary?: string; displayTitle?: string; displayKind?: string },
  verbosity: "low" | "medium" | "high" = "medium",
): string {
  return formatToolCall(tool, verbosity);
}

export function formatPlan(entries: { content: string; status: string }[], verbosity: "low" | "medium" | "high" = "medium"): string {
  if (verbosity === "medium") {
    const done = entries.filter((e) => e.status === "completed").length;
    return `📋 **Plan:** ${done}/${entries.length} steps completed`;
  }
  const statusIconMap: Record<string, string> = {
    pending: "⏳", in_progress: "🔄", completed: "✅",
  };
  const lines = entries.map((e, i) => `${statusIconMap[e.status] || "⬜"} ${i + 1}. ${e.content}`);
  return `**Plan:**\n${lines.join("\n")}`;
}

export function formatUsage(
  usage: { tokensUsed?: number; contextSize?: number; cost?: number },
  verbosity: "low" | "medium" | "high" = "medium",
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

// Usage card renderer for Adaptive Cards
interface UsageData {
  tokensUsed?: number;
  contextSize?: number;
  cost?: number;
  duration?: number;
}

export function renderUsageCard(usage: UsageData, mode: "low" | "medium" | "high"): { body: unknown[] } {
  const text = formatUsage(usage, mode);
  return { body: [{ type: "TextBlock", text, wrap: true }] };
}