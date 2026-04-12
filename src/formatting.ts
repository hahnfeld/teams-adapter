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

/**
 * Split text into chunks that fit within maxLength, preferring paragraph
 * and line boundaries. Handles single lines longer than maxLength by
 * hard-splitting at the limit.
 */
export function splitMessage(text: string, maxLength: number): string[] {
  if (text.length <= maxLength) return [text];
  const paragraphs = text.split("\n\n");
  const chunks: string[] = [];
  let current = "";

  const pushCurrent = () => {
    if (current) {
      chunks.push(current);
      current = "";
    }
  };

  // Hard-split a single string that exceeds maxLength
  const hardSplit = (s: string) => {
    while (s.length > maxLength) {
      chunks.push(s.slice(0, maxLength));
      s = s.slice(maxLength);
    }
    return s; // remainder
  };

  for (const para of paragraphs) {
    const candidate = current ? `${current}\n\n${para}` : para;

    if (candidate.length <= maxLength) {
      current = candidate;
      continue;
    }

    // Candidate exceeds limit — flush current and process para separately
    pushCurrent();

    if (para.length <= maxLength) {
      current = para;
      continue;
    }

    // Para exceeds limit — try splitting on line boundaries
    const lines = para.split("\n");
    for (const line of lines) {
      if (line.length > maxLength) {
        // Line itself exceeds limit — hard-split it
        pushCurrent();
        current = hardSplit(line);
        continue;
      }

      const lineCandidate = current ? `${current}\n${line}` : line;
      if (lineCandidate.length > maxLength) {
        pushCurrent();
        current = line;
      } else {
        current = lineCandidate;
      }
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
  return text.replace(/^```[\w]*\n?/gm, "").replace(/\n?```$/gm, "").trim();
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

// ─── Adaptive Card builders for Teams-native rendering ──────────────────

interface UsageData {
  tokensUsed?: number;
  contextSize?: number;
  cost?: number;
  duration?: number;
}

/** Structured usage card with FactSet and visual progress indicator */
export function renderUsageCard(usage: UsageData, mode: "low" | "medium" | "high"): { body: unknown[] } {
  const { tokensUsed, contextSize, cost, duration } = usage;

  if (tokensUsed == null) {
    return { body: [{ type: "TextBlock", text: "📊 Usage data unavailable", wrap: true, isSubtle: true }] };
  }

  if (mode === "low") {
    return { body: [{ type: "TextBlock", text: `📊 ${formatTokens(tokensUsed)} tokens`, wrap: true }] };
  }

  const body: unknown[] = [];

  // Header
  const pct = contextSize ? Math.round((tokensUsed / contextSize) * 100) : null;
  const headerColor = pct != null && pct >= 85 ? "Attention" : "Default";
  body.push({ type: "TextBlock", text: "📊 Usage Summary", weight: "Bolder", size: "Medium", color: headerColor });

  // Facts
  const facts: Array<{ title: string; value: string }> = [
    { title: "Tokens", value: formatTokens(tokensUsed) },
  ];
  if (contextSize) facts.push({ title: "Context", value: `${formatTokens(contextSize)} (${pct}%)` });
  if (cost != null) facts.push({ title: "Cost", value: `$${cost.toFixed(4)}` });
  if (duration != null) facts.push({ title: "Duration", value: `${(duration / 1000).toFixed(1)}s` });
  body.push({ type: "FactSet", facts, separator: true });

  // Visual progress bar (context usage)
  if (contextSize && mode === "high") {
    body.push({
      type: "TextBlock",
      text: `\`${progressBar(tokensUsed / contextSize, 20)}\` ${pct}%`,
      fontType: "Monospace",
      size: "Small",
      isSubtle: true,
    });
  }

  return { body };
}

interface ToolCallCardMeta {
  id: string;
  name: string;
  kind?: string;
  status?: string;
  rawInput?: unknown;
  content?: unknown;
  displaySummary?: string;
  displayTitle?: string;
  displayKind?: string;
  viewerLinks?: { file?: string; diff?: string };
  viewerFilePath?: string;
}

/** Structured tool call card with icon, name, status, and optional details */
export function renderToolCallCard(tool: ToolCallCardMeta, verbosity: "low" | "medium" | "high"): { body: unknown[]; actions?: unknown[] } {
  const icon = resolveToolIcon(tool.kind ?? "", tool.displayKind, tool.status);
  const name = tool.name || "Tool";
  const label = verbosity === "low"
    ? formatToolTitle(name, tool.rawInput, tool.displayTitle)
    : formatToolSummary(name, tool.rawInput, tool.displaySummary);

  const statusColor = tool.status === "completed" ? "Good" : tool.status === "error" ? "Attention" : "Default";

  const body: unknown[] = [];

  // Header row: icon + tool name + status
  body.push({
    type: "ColumnSet",
    columns: [
      {
        type: "Column",
        width: "auto",
        items: [{ type: "TextBlock", text: icon, size: "Medium" }],
      },
      {
        type: "Column",
        width: "stretch",
        items: [
          { type: "TextBlock", text: `**${label}**`, wrap: true },
          ...(tool.status ? [{ type: "TextBlock", text: tool.status, size: "Small", isSubtle: true, color: statusColor, spacing: "None" }] : []),
        ],
      },
    ],
  });

  // High verbosity: show input/output
  if (verbosity === "high") {
    const maxLen = 400;
    if (tool.rawInput) {
      const inputStr = typeof tool.rawInput === "string" ? tool.rawInput : JSON.stringify(tool.rawInput, null, 2);
      if (inputStr && inputStr !== "{}") {
        body.push({ type: "TextBlock", text: "**Input:**", size: "Small", spacing: "Medium" });
        body.push({ type: "TextBlock", text: `\`\`\`\n${truncateContent(inputStr, maxLen)}\n\`\`\``, wrap: true, fontType: "Monospace", size: "Small" });
      }
    }
    const contentText = stripCodeFences(extractContentText(tool.content));
    if (contentText) {
      body.push({ type: "TextBlock", text: "**Output:**", size: "Small", spacing: "Medium" });
      body.push({ type: "TextBlock", text: `\`\`\`\n${truncateContent(contentText, maxLen)}\n\`\`\``, wrap: true, fontType: "Monospace", size: "Small" });
    }
  }

  // Actions: viewer links
  const actions: unknown[] = [];
  if (tool.viewerLinks) {
    const fn = tool.viewerFilePath?.split("/").pop() || "file";
    if (tool.viewerLinks.file) {
      actions.push({ type: "Action.OpenUrl", title: `View ${fn}`, url: tool.viewerLinks.file });
    }
    if (tool.viewerLinks.diff) {
      actions.push({ type: "Action.OpenUrl", title: `View diff`, url: tool.viewerLinks.diff });
    }
  }

  return { body, ...(actions.length > 0 ? { actions } : {}) };
}

/** Structured plan card with status-colored rows */
export function renderPlanCard(entries: { content: string; status: string }[], verbosity: "low" | "medium" | "high"): { body: unknown[] } {
  const done = entries.filter((e) => e.status === "completed").length;
  const total = entries.length;

  const body: unknown[] = [];

  // Header with progress
  body.push({
    type: "ColumnSet",
    columns: [
      { type: "Column", width: "stretch", items: [{ type: "TextBlock", text: "📋 **Plan**", weight: "Bolder" }] },
      { type: "Column", width: "auto", items: [{ type: "TextBlock", text: `${done}/${total}`, isSubtle: true }] },
    ],
  });

  if (verbosity === "medium" && total > 5) {
    // Compact summary for medium verbosity with many steps
    body.push({ type: "TextBlock", text: `${done} of ${total} steps completed`, isSubtle: true });
    return { body };
  }

  // Step rows with status colors
  const statusColors: Record<string, string> = {
    completed: "Good", in_progress: "Accent", pending: "Default",
  };
  const statusIcons: Record<string, string> = {
    completed: "✅", in_progress: "🔄", pending: "⬜",
  };

  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    const icon = statusIcons[e.status] || "⬜";
    const color = statusColors[e.status] || "Default";
    body.push({
      type: "ColumnSet",
      spacing: "Small",
      columns: [
        { type: "Column", width: "auto", items: [{ type: "TextBlock", text: icon }] },
        { type: "Column", width: "stretch", items: [{ type: "TextBlock", text: `${i + 1}. ${e.content}`, wrap: true, color }] },
      ],
    });
  }

  return { body };
}

// ─── Citation entities for Teams AI content ─────────────────────────────

interface CitationSource {
  /** File path or tool name */
  name: string;
  /** URL to view the file or diff */
  url: string;
  /** Short description or content snippet */
  abstract?: string;
  /** Icon type: "Source Code", "PDF", "Image", etc. */
  iconType?: string;
}

/**
 * Build Teams citation entities from file references.
 *
 * Teams displays citations as numbered in-text references [1] with hover popups
 * showing the source name, description, and link. Max 20 per message.
 *
 * @see https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/bot-messages-ai-generated-content
 */
export function buildCitationEntities(sources: CitationSource[]): unknown[] {
  if (sources.length === 0) return [];

  return [{
    type: "https://schema.org/Message",
    "@type": "Message",
    "@context": "https://schema.org",
    additionalType: ["AIGeneratedContent"],
    citation: sources.slice(0, 20).map((source, i) => ({
      "@type": "Claim",
      position: i + 1,
      appearance: {
        "@type": "DigitalDocument",
        name: source.name.slice(0, 80),
        url: source.url,
        ...(source.abstract ? { abstract: source.abstract.slice(0, 160) } : {}),
        image: { "@type": "ImageObject", name: source.iconType ?? guessIconType(source.name) },
      },
    })),
  }];
}

/**
 */

/** Guess the Teams citation icon type from a file extension */
function guessIconType(name: string): string {
  const ext = name.split(".").pop()?.toLowerCase() ?? "";
  const codeExts = new Set(["ts", "js", "tsx", "jsx", "py", "rs", "go", "java", "c", "cpp", "h", "cs", "rb", "swift", "kt", "sh", "yaml", "yml", "json", "toml", "xml", "html", "css", "scss", "sql", "md"]);
  if (codeExts.has(ext)) return "Source Code";
  if (ext === "pdf") return "PDF";
  if (["png", "jpg", "jpeg", "gif", "svg", "webp"].includes(ext)) return "Image";
  if (["doc", "docx"].includes(ext)) return "Microsoft Word";
  if (["xls", "xlsx"].includes(ext)) return "Microsoft Excel";
  if (["ppt", "pptx"].includes(ext)) return "Microsoft PowerPoint";
  return "Source Code";
}