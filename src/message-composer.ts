/**
 * Message composer — manages a single "main message" per session in Teams.
 *
 * Emulates Claude Code's terminal UX using Adaptive Cards:
 *   - Title: bold, persistent — session name
 *   - Thinking blocks: italic, 💭 prefix, ALWAYS append (never replaced)
 *   - Tool progress: 🔄 Running... with elapsed time, persistent as historical record
 *   - Tool result: 📄 Result... appears AFTER progress (both coexist)
 *   - Streaming text: root-level, appended
 *   - Usage: italic footer, persistent
 *
 * Tool output during execution is indented as children under the tool block.
 * When a tool completes, a NEW tool-result entry is appended — the progress
 * entry stays as a historical record (matching Claude Code's behavior).
 */
import type { TurnContext } from "@microsoft/agents-hosting";
import { log } from "@openacp/plugin-sdk";
import { CardFactory } from "@microsoft/agents-hosting";
import type { ConversationRateLimiter } from "./rate-limiter.js";

// ─── Re-exports for backwards compat ───────────────────────────────────────────

export interface MessageRef {
  activityId: string;
  conversationId: string;
  serviceUrl: string;
}

export type AcquireBotToken = () => Promise<string | null>;

const MAX_ROOT_TEXT_LENGTH = 25_000;
/** How long to wait with no activity before warning about truncation. */
const STALL_TIMEOUT = 120_000;

// ─── Brand Colors ────────────────────────────────────────────────────────────
// Primary blue #01426a | Dark blue #00274a | Purple #463c8f
// Cyan #5de3f7 | Green #a3cd6a | Magenta #ce0c88
// Text default #2a2a2a | Muted #676767 | White #ffffff

const BRAND = {
  blue: "#01426a",
  blueDark: "#00274a",
  purple: "#463c8f",
  cyan: "#5de3f7",
  green: "#a3cd6a",
  magenta: "#ce0c88",
  textDefault: "#2a2a2a",
  textMuted: "#676767",
  white: "#ffffff",
  surfaceMuted: "#f7f7f7",
} as const;

// ─── Entry Types ───────────────────────────────────────────────────────────────

type BodyEntry =
  | { id: string; kind: "title"; text: string }
  | { id: string; kind: "tool"; toolName: string; startedAt: number; result?: string; endedAt?: number; children: TextChild[] }
  | { id: string; kind: "text"; text: string }
  | { id: string; kind: "thought"; text: string }
  | { id: string; kind: "usage"; text: string }
  | { id: string; kind: "divider" };

/** Text content inside a tool block — appended as tool runs */
type TextChild = { text: string };

// ─── ID generation ─────────────────────────────────────────────────────────────

let _idCounter = 0;
function nextId(): string {
  return `e${++_idCounter}_${Date.now().toString(36)}`;
}

// ─── Adaptive Card Builder ────────────────────────────────────────────────────

function buildCardBody(entries: BodyEntry[]): unknown[] {
  const blocks: unknown[] = [];
  const now = Date.now();

  for (const entry of entries) {
    switch (entry.kind) {
      case "title":
        blocks.push({
          type: "TextBlock",
          text: `**${escapeMd(entry.text)}**`,
          weight: "Bolder",
          size: "Medium",
          fontType: "Monospace",
          spacing: "None",
        });
        break;

      case "tool": {
        const elapsed = entry.result
          ? formatElapsed((entry.endedAt ?? entry.startedAt) - entry.startedAt)
          : formatElapsed(now - entry.startedAt);
        blocks.push({
          type: "Container",
          items: [
            {
              type: "TextBlock",
              text: entry.result
                ? `🔧 ${escapeMd(entry.toolName)}`
                : `🔧 ${escapeMd(entry.toolName)}…  (${elapsed})`,
              size: "Small",
              fontType: "Monospace",
              spacing: "None",
            },
            // Result: indented with L-shaped bar
            ...(entry.result
              ? [{
                  type: "TextBlock" as const,
                  text: `\u00A0\u00A0⎿ ${escapeMd(entry.result)}  (${elapsed})`,
                  size: "Small" as const,
                  fontType: "Monospace",
                  wrap: true,
                  spacing: "None",
                }]
              : []),
            // Children: further indented
            ...entry.children.map((c) => ({
              type: "TextBlock",
              text: `\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0${c.text}`,
              size: "Small",
              fontType: "Monospace",
              wrap: true,
              spacing: "None",
            })),
          ],
          padding: "Small",
          width: "stretch",
        });
        break;
      }

      case "text":
        blocks.push({
          type: "TextBlock",
          text: entry.text,
          size: "Small",
          fontType: "Monospace",
          wrap: true,
          spacing: "None",
        });
        break;

      case "thought":
        blocks.push({
          type: "TextBlock",
          text: entry.text,
          italic: true,
          size: "Small",
          fontType: "Monospace",
          wrap: true,
          spacing: "None",
        });
        break;

      case "usage":
        blocks.push({
          type: "TextBlock",
          text: `*${escapeMd(entry.text)}*`,
          italic: true,
          size: "Small",
          fontType: "Monospace",
          spacing: "None",
        });
        break;

      case "divider":
        blocks.push({
          type: "TextBlock",
          text: "─".repeat(30),
          size: "Small",
          fontType: "Monospace",
          spacing: "None",
        });
        break;
    }
  }

  return blocks;
}

function formatElapsed(ms: number): string {
  if (ms < 1000) return `${ms}ms`;
  if (ms < 60_000) return `${(ms / 1000).toFixed(1)}s`;
  return `${Math.floor(ms / 60_000)}m ${Math.floor((ms % 60_000) / 1000)}s`;
}

function escapeMd(text: string): string {
  return text.replace(/\[/g, "\\[").replace(/\*/g, "\\*").replace(/\]/g, "\\]");
}

// ─── SessionMessage ────────────────────────────────────────────────────────────

export class SessionMessage {
  private entries: BodyEntry[] = [];
  private titleId: string | null = null;
  private usageId: string | null = null;
  /** The active tool whose children accumulate streaming output */
  private toolActive: string | null = null;
  private ref: MessageRef | null = null;
  private lastSent = "";
  private stallTimer?: ReturnType<typeof setTimeout>;
  /** Interval handle for periodic elapsed-time updates on running tools */
  private tickInterval?: ReturnType<typeof setInterval>;

  constructor(
    private context: TurnContext,
    private conversationId: string,
    private sessionId: string,
    private rateLimiter: ConversationRateLimiter,
    private acquireBotToken: AcquireBotToken,
  ) {}

  updateContext(context: TurnContext): void {
    this.context = context;
  }

  getRef(): MessageRef | null {
    return this.ref;
  }

  getBody(): string {
    return this.entries
      .filter((e) => e.kind === "text")
      .map((e) => (e as { kind: "text"; text: string }).text)
      .join("");
  }

  getFooter(): string | null {
    const usage = this.entries.find((e) => e.kind === "usage");
    return usage ? (usage as { kind: "usage"; text: string }).text : null;
  }

  // ─── Entry API ───────────────────────────────────────────────────────────

  /** Set the persistent session title (bold, survives finalize). */
  setTitle(text: string): void {
    if (this.titleId) {
      const entry = this.findEntry(this.titleId);
      if (entry && entry.kind === "title") entry.text = text;
    } else {
      this.titleId = nextId();
      this.entries.unshift({ id: this.titleId, kind: "title", text });
    }
    this.requestFlush();
  }

  /**
   * Add a tool-progress entry (🔄 Running...). Sets toolActive so subsequent
   * addText() calls route children here. Returns entry id for tracking.
   */
  addToolStart(toolName: string, _params?: string): string {
    const id = nextId();
    const startedAt = Date.now();
    this.entries.push({ id, kind: "tool", toolName, startedAt, children: [] });
    this.toolActive = id;
    this.resetStallTimer();
    this.startTickInterval();
    this.requestFlush();
    return id;
  }

  /**
   * Append result to the tool entry (transforms it from running to complete).
   * Subsequent text goes to the same entry's children.
   */
  addToolResult(id: string, result: string): void {
    const entry = this.entries.find((e) => e.id === id);
    if (!entry || entry.kind !== "tool") {
      // Entry not found — create a standalone tool entry
      const startedAt = Date.now();
      this.entries.push({ id: nextId(), kind: "tool", toolName: "", startedAt, result, endedAt: startedAt, children: [] });
      this.toolActive = null;
      this.stopTickInterval();
      this.requestFlush();
      return;
    }

    entry.result = result;
    entry.endedAt = Date.now();
    this.toolActive = null; // Subsequent text goes to root level
    this.stopTickInterval();
    this.requestFlush();
  }

  /** Add text — goes to toolActive children if a tool is running, else root text entry. */
  addText(text: string): void {
    if (!text) return;

    if (this.toolActive) {
      const entry = this.findEntry(this.toolActive);
      if (entry && entry.kind === "tool") {
        entry.children.push({ text });
        this.resetStallTimer();
        this.requestFlush();
        // Check for overflow (root text limit)
        if (this.ref) this.checkSplit();
        return;
      }
    }

    // Root-level text — append to last root text entry or create new
    const lastText = this.entries.filter((e) => e.kind === "text").at(-1);
    if (lastText) {
      lastText.text += text;
    } else {
      this.entries.push({ id: nextId(), kind: "text", text });
    }
    this.resetStallTimer();
    this.requestFlush();
    // Check for overflow after root text grows
    if (this.ref) this.checkSplit();
  }

  /**
   * Append text to the last thought entry, or create a new one if none exists.
   * If text starts with "Thinking..." and it's a fresh entry, wrap with newlines.
   * Subsequent chunks are appended as plain text with no bubble prefix.
   */
  addThought(text: string): void {
    const thoughts = this.entries.filter((e) => e.kind === "thought");
    const last = thoughts[thoughts.length - 1];
    if (last) {
      // Add a space between consecutive chunks so they don't run together
      last.text += ` ${text}`;
    } else {
      this.entries.push({ id: nextId(), kind: "thought", text });
    }
    this.requestFlush();
  }

  /** Set or replace the usage entry (only one at a time). */
  setUsage(text: string): void {
    if (this.usageId) {
      const entry = this.findEntry(this.usageId);
      if (entry && entry.kind === "usage") entry.text = text;
    } else {
      this.usageId = nextId();
      this.entries.push({ id: this.usageId, kind: "usage", text });
    }
    this.requestFlush();
  }

  /** Add a divider entry. */
  appendDivider(): void {
    this.entries.push({ id: nextId(), kind: "divider" });
    this.requestFlush();
  }

  // ─── Entry helpers ────────────────────────────────────────────────────────

  private findEntry(id: string): BodyEntry | undefined {
    for (const entry of this.entries) {
      if (entry.id === id) return entry;
    }
    return undefined;
  }

  private removeEntry(id: string): void {
    this.entries = this.entries.filter((e) => e.id !== id);
  }

  // ─── Periodic tick for elapsed time updates ───────────────────────────────

  /**
   * Start a 1-second interval that updates elapsed time on running tool-progress
   * entries. Called when a tool starts, stopped when tool completes or finalize.
   */
  private startTickInterval(): void {
    if (this.tickInterval) return;
    this.tickInterval = setInterval(() => {
      // Only tick if there's an active tool-progress entry
      const hasRunningTool = this.entries.some((e) => e.kind === "tool" && !e.result);
      if (!hasRunningTool) {
        this.stopTickInterval();
        return;
      }
      this.requestFlush();
    }, 1_000);
    if (this.tickInterval.unref) this.tickInterval.unref();
  }

  private stopTickInterval(): void {
    if (this.tickInterval) {
      clearInterval(this.tickInterval);
      this.tickInterval = undefined;
    }
  }

  // ─── Card building ─────────────────────────────────────────────────────────

  private buildCard(): Record<string, unknown> {
    const body = buildCardBody(this.entries);

    return {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        ...(body.length > 0 ? body : [{ type: "TextBlock", text: "…" }]),
      ],
      // Use full available width
      width: "stretch",
    };
  }

  // ─── Flush / Rate limiting ─────────────────────────────────────────────────

  private flushTimer?: ReturnType<typeof setTimeout>;

  private requestFlush(): void {
    if (this.flushTimer) clearTimeout(this.flushTimer);
    this.flushTimer = setTimeout(() => {
      this.flushTimer = undefined;
      const card = this.buildCard();
      const key = this.ref ? `update:${this.ref.activityId}` : `new:${this.sessionId}`;
      this.rateLimiter.enqueue(
        this.conversationId,
        () => this.flush(),
        key,
      ).catch((err) => {
        log.warn({ err, sessionId: this.sessionId }, "[SessionMessage] flush failed");
      });
    }, 500);
  }

  private async flush(): Promise<void> {
    const card = this.buildCard();
    const cardStr = JSON.stringify(card);
    if (!cardStr || cardStr === this.lastSent) return;

    if (!this.ref) {
      const result = await sendCard(this.context, card) as { id?: string } | undefined;
      if (result?.id) {
        this.ref = {
          activityId: result.id,
          conversationId: this.context.activity.conversation?.id as string,
          serviceUrl: this.context.activity.serviceUrl as string,
        };
      }
      this.lastSent = cardStr;
    } else {
      const success = await this.updateCardViaRest(card);
      if (success) {
        this.lastSent = cardStr;
      }
    }

    // Content arrived while flushing — request another flush
    const current = this.buildCard();
    if (JSON.stringify(current) !== this.lastSent) {
      this.requestFlush();
    }
  }

  private async updateCardViaRest(card: Record<string, unknown>): Promise<boolean> {
    if (!this.ref) return false;
    const token = await this.acquireBotToken();
    if (!token) return false;

    const url = `${this.ref.serviceUrl}/v3/conversations/${encodeURIComponent(this.ref.conversationId)}/activities/${encodeURIComponent(this.ref.activityId)}`;

    try {
      const response = await fetch(url, {
        method: "PUT",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${token}`,
        },
        body: JSON.stringify({
          type: "message",
          attachments: [CardFactory.adaptiveCard(card)],
        }),
      });

      if (!response.ok) {
        log.warn({ status: response.status, sessionId: this.sessionId }, "[SessionMessage] REST card update failed");
        return false;
      }
      return true;
    } catch (err) {
      log.warn({ err, sessionId: this.sessionId }, "[SessionMessage] REST card update error");
      return false;
    }
  }

  // ─── Stall timer ───────────────────────────────────────────────────────────

  private resetStallTimer(): void {
    if (this.stallTimer) clearTimeout(this.stallTimer);
    this.stallTimer = setTimeout(() => {
      const hasContent = this.entries.some(
        (e) => (e.kind === "text" && e.text.length > 0) ||
               (e.kind === "tool" && e.children.length > 0),
      );
      const hasUsage = this.entries.some((e) => e.kind === "usage");
      if (hasContent && !hasUsage) {
        log.warn({ sessionId: this.sessionId }, "[SessionMessage] Stream stalled — adding cutoff notice");
        this.entries.push({ id: nextId(), kind: "divider" });
        this.entries.push({
          id: nextId(),
          kind: "text",
          text: "\n\n---\n_Response was cut short — the model likely reached its output token limit. Send a follow-up message to continue._",
        });
        this.requestFlush();
      }
    }, STALL_TIMEOUT);
    if (this.stallTimer.unref) this.stallTimer.unref();
  }

  // ─── Finalize ─────────────────────────────────────────────────────────────

  async finalize(): Promise<MessageRef | null> {
    this.stopTickInterval();
    if (this.stallTimer) {
      clearTimeout(this.stallTimer);
      this.stallTimer = undefined;
    }

    const card = this.buildCard();
    const cardStr = JSON.stringify(card);
    if (cardStr && cardStr !== this.lastSent) {
      if (!this.ref) {
        const result = await sendCard(this.context, card) as { id?: string } | undefined;
        if (result?.id) {
          this.ref = {
            activityId: result.id,
            conversationId: this.context.activity.conversation?.id as string,
            serviceUrl: this.context.activity.serviceUrl as string,
          };
        }
      } else {
        await this.updateCardViaRest(card);
      }
      this.lastSent = cardStr;
    }

    return this.ref;
  }

  /** Legacy compat — strip pattern from root text entries. */
  async stripPattern(pattern: RegExp): Promise<void> {
    for (const entry of this.entries) {
      if (entry.kind === "text") {
        try {
          entry.text = entry.text.replace(pattern, "").trim();
        } catch { /* leave unchanged */ }
      }
    }
  }

  /**
   * Check if root text exceeds limit and split into a new message.
   * Called from addText when a message is already sent (this.ref != null).
   */
  private checkSplit(): void {
    const rootText = this.entries
      .filter((e) => e.kind === "text")
      .map((e) => (e as { kind: "text"; text: string }).text)
      .join("");

    if (rootText.length <= MAX_ROOT_TEXT_LENGTH) return;
    this.split();
  }

  /** For split: finalize current message at limit, start fresh. */
  private split(): void {
    const rootText = this.entries
      .filter((e) => e.kind === "text")
      .map((e) => (e as { kind: "text"; text: string }).text)
      .join("");

    if (rootText.length <= MAX_ROOT_TEXT_LENGTH) return;

    let accLen = 0;
    const entriesToKeep: BodyEntry[] = [];
    const entriesToOverflow: BodyEntry[] = [];

    for (const entry of this.entries) {
      if (entry.kind === "text") {
        const text = (entry as { kind: "text"; text: string }).text;
        if (accLen + text.length <= MAX_ROOT_TEXT_LENGTH) {
          entriesToKeep.push(entry);
          accLen += text.length;
        } else {
          entriesToOverflow.push(entry);
        }
      } else {
        entriesToKeep.push(entry);
      }
    }

    log.info({ sessionId: this.sessionId, kept: entriesToKeep.length, overflow: entriesToOverflow.length }, "[SessionMessage] splitting");

    this.entries = entriesToKeep;
    if (entriesToOverflow.length > 0) {
      const overflowText = entriesToOverflow
        .filter((e) => e.kind === "text")
        .map((e) => (e as { kind: "text"; text: string }).text)
        .join("");
      if (overflowText) {
        this.entries.push({ id: nextId(), kind: "text", text: overflowText });
      }
    }

    if (this.ref) {
      const card = this.buildCard();
      this.rateLimiter.enqueue(
        this.conversationId,
        () => this.updateCardViaRest(card).then(() => {}),
        `update:${this.ref.activityId}`,
      ).catch(() => {});
    }

    this.ref = null;
    this.lastSent = "";
    this.requestFlush();
  }

  // ─── Legacy API (no-ops for compat) ────────────────────────────────────────

  /** @deprecated — use addThought instead */
  setHeader(_text: string): void { /* no-op */ }

  /** @deprecated — use addToolResult instead */
  setHeaderResult(_text: string): void { /* no-op */ }

  /** @deprecated — use addText instead */
  appendBody(text: string): void {
    this.addText(text);
  }

  /** @deprecated — use setUsage instead */
  setFooter(text: string): void {
    this.setUsage(text);
  }

  /** @deprecated — use setUsage instead */
  appendFooter(text: string): void {
    const current = this.getFooter();
    this.setUsage(current ? `${current} · ${text}` : text);
  }

  /** @deprecated — header zone is gone */
  clearHeader(): void { /* no-op */ }

  /** @deprecated — thoughts persist, use addThought */
  removeThought(): void { /* no-op */ }

  /** @deprecated — use addToolStart + addToolResult */
  updateToolResult(_id: string, _result: string): void { /* no-op */ }
}

// ─── SessionMessageManager ─────────────────────────────────────────────────────

export class SessionMessageManager {
  private messages = new Map<string, SessionMessage>();
  private planRefs = new Map<string, MessageRef>();

  constructor(
    private rateLimiter: ConversationRateLimiter,
    private acquireBotToken: AcquireBotToken,
  ) {}

  getOrCreate(sessionId: string, context: TurnContext): SessionMessage {
    let msg = this.messages.get(sessionId);
    if (!msg) {
      const conversationId = context.activity.conversation?.id as string;
      msg = new SessionMessage(context, conversationId, sessionId, this.rateLimiter, this.acquireBotToken);
      this.messages.set(sessionId, msg);
    } else {
      msg.updateContext(context);
    }
    return msg;
  }

  get(sessionId: string): SessionMessage | undefined {
    return this.messages.get(sessionId);
  }

  has(sessionId: string): boolean {
    return this.messages.has(sessionId);
  }

  async finalize(sessionId: string): Promise<MessageRef | null> {
    const msg = this.messages.get(sessionId);
    if (!msg) return null;
    this.messages.delete(sessionId);
    this.planRefs.delete(sessionId);
    return msg.finalize();
  }

  getPlanRef(sessionId: string): MessageRef | undefined {
    return this.planRefs.get(sessionId);
  }

  setPlanRef(sessionId: string, ref: MessageRef): void {
    this.planRefs.set(sessionId, ref);
  }

  cleanup(sessionId: string): void {
    const msg = this.messages.get(sessionId);
    if (msg) {
      msg.finalize().catch(() => {});
    }
    this.messages.delete(sessionId);
    this.planRefs.delete(sessionId);
  }
}

// ─── Helpers ───────────────────────────────────────────────────────────────────

async function sendCard(context: TurnContext, card: Record<string, unknown>): Promise<unknown> {
  const activity = {
    type: "message",
    attachments: [CardFactory.adaptiveCard(card)],
  };
  if (typeof (context as any).send === "function") {
    return (context as any).send(activity);
  }
  return (context.sendActivity as Function)(activity);
}