/**
 * Message composer — manages a single "main message" per session in Teams.
 *
 * Uses an Adaptive Card with a hierarchical entry model:
 *   - title: bold, persistent — session name, survives finalize
 *   - tool-start/tool-result: shows tool name + status, with child text entries
 *   - text: streaming text at root level
 *   - thought: italic, can be replaced/removed
 *   - usage: appended at end
 *
 * Text output during a tool call is indented as a child of that tool entry.
 * This matches Claude Code's visual model where tool results "fill in" under
 * the tool block before being replaced by the final result.
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
/** How long to wait with no activity (text chunks or tool changes) before warning about truncation. */
const STALL_TIMEOUT = 120_000;

// ─── Entry Types ───────────────────────────────────────────────────────────────

type BodyEntry =
  | { id: string; kind: "title"; text: string }
  | { id: string; kind: "tool-start"; toolName: string; params?: string; status: "running"; children: TextChild[] }
  | { id: string; kind: "tool-result"; toolName: string; result: string; children: TextChild[] }
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

  for (const entry of entries) {
    switch (entry.kind) {
      case "title":
        blocks.push({
          type: "TextBlock",
          text: `**${escapeMd(entry.text)}**`,
          weight: "Bolder",
          size: "Medium",
          spacing: "None",
        });
        break;

      case "tool-start":
        blocks.push({
          type: "Container",
          items: [
            {
              type: "TextBlock",
              text: `🔄 ${escapeMd(entry.toolName)}`,
              isSubtle: false,
              weight: "Bolder",
              size: "Small",
              spacing: "None",
            },
            ...entry.children.map((c) => ({
              type: "TextBlock",
              text: `    ${c.text}`,
              size: "Small",
              isSubtle: true,
              spacing: "None",
            })),
          ],
          style: "emphasis",
          padding: "Small",
        });
        break;

      case "tool-result":
        blocks.push({
          type: "Container",
          items: [
            {
              type: "TextBlock",
              text: `📄 ${escapeMd(entry.result)}`,
              weight: "Bolder",
              size: "Small",
              spacing: "None",
            },
            ...entry.children.map((c) => ({
              type: "TextBlock",
              text: `    ${c.text}`,
              size: "Small",
              isSubtle: false,
              spacing: "None",
            })),
          ],
          style: "good",
          padding: "Small",
        });
        break;

      case "text":
        blocks.push({
          type: "TextBlock",
          text: entry.text,
          size: "Small",
          wrap: true,
          spacing: "None",
        });
        break;

      case "thought":
        blocks.push({
          type: "TextBlock",
          text: `💭 ${entry.text}`,
          italic: true,
          size: "Small",
          isSubtle: true,
          spacing: "None",
        });
        break;

      case "usage":
        blocks.push({
          type: "TextBlock",
          text: `*${escapeMd(entry.text)}*`,
          italic: true,
          size: "Small",
          isSubtle: true,
          spacing: "None",
        });
        break;

      case "divider":
        blocks.push({
          type: "TextBlock",
          text: "─".repeat(20),
          size: "Small",
          isSubtle: true,
          spacing: "None",
        });
        break;
    }
  }

  return blocks;
}

function escapeMd(text: string): string {
  // Escape markdown special chars that would render incorrectly in Adaptive Cards
  return text.replace(/\[/g, "\\[").replace(/\*/g, "\\*").replace(/\]/g, "\\]");
}

// ─── SessionMessage ────────────────────────────────────────────────────────────

/**
 * A single main message with hierarchical entry structure.
 * Entries are rendered into an Adaptive Card on each flush.
 */
export class SessionMessage {
  private entries: BodyEntry[] = [];
  private titleId: string | null = null;
  private usageId: string | null = null;
  private thoughtId: string | null = null;
  private toolActive: string | null = null;
  private ref: MessageRef | null = null;
  private lastSent = "";
  private stallTimer?: ReturnType<typeof setTimeout>;

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
    // Legacy compat — returns concatenated text entries
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

  /** Add a new tool-start entry, returns entry id. Sets toolActive. */
  addToolStart(toolName: string, params?: string): string {
    const id = nextId();
    this.entries.push({ id, kind: "tool-start", toolName, params, status: "running", children: [] });
    this.toolActive = id;
    this.resetStallTimer();
    this.requestFlush();
    return id;
  }

  /**
   * Replace a tool-start entry with a tool-result at the same id.
   * Clears toolActive. Preserves any child text that streamed during execution.
   */
  updateToolResult(id: string, result: string): void {
    const idx = this.entries.findIndex((e) => e.id === id);
    if (idx === -1) return;

    const entry = this.entries[idx];
    if (entry.kind !== "tool-start") return;

    const children = [...entry.children];
    this.entries[idx] = { id, kind: "tool-result", toolName: entry.toolName, result, children };
    if (this.toolActive === id) this.toolActive = null;
    this.requestFlush();
  }

  /** Add text — goes to toolActive children if a tool is running, else root text entry. */
  addText(text: string): void {
    if (!text) return;

    if (this.toolActive) {
      const entry = this.findEntry(this.toolActive);
      if (entry && (entry.kind === "tool-start" || entry.kind === "tool-result")) {
        entry.children.push({ text });
        this.resetStallTimer();
        this.requestFlush();
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
  }

  /** Add or replace the thought entry (only one at a time). */
  addOrReplaceThought(text: string): void {
    if (this.thoughtId) {
      const entry = this.findEntry(this.thoughtId);
      if (entry && entry.kind === "thought") {
        entry.text = text;
      }
    } else {
      this.thoughtId = nextId();
      this.entries.push({ id: this.thoughtId, kind: "thought", text });
    }
    this.requestFlush();
  }

  /** Remove the thought entry (e.g., when not relevant to final result). */
  removeThought(): void {
    if (!this.thoughtId) return;
    this.removeEntry(this.thoughtId);
    this.thoughtId = null;
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
      if (entry.kind === "tool-start" || entry.kind === "tool-result") {
        // Children are flat TextChild[], not full BodyEntry — no nested search needed
      }
    }
    return undefined;
  }

  private updateEntry(id: string, patch: Partial<BodyEntry>): void {
    const idx = this.entries.findIndex((e) => e.id === id);
    if (idx !== -1) {
      this.entries[idx] = { ...this.entries[idx], ...patch } as BodyEntry;
    }
  }

  private removeEntry(id: string): void {
    this.entries = this.entries.filter((e) => e.id !== id);
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
    };
  }

  // ─── Flush / Rate limiting ─────────────────────────────────────────────────

  /** Request a flush through the rate limiter. */
  private requestFlush(): void {
    const card = this.buildCard();
    const key = this.ref ? `update:${this.ref.activityId}` : `new:${this.sessionId}`;
    this.rateLimiter.enqueue(
      this.conversationId,
      () => this.flush(),
      key,
    ).catch((err) => {
      log.warn({ err, sessionId: this.sessionId }, "[SessionMessage] flush failed");
    });
  }

  private async flush(): Promise<void> {
    const card = this.buildCard();
    const cardStr = JSON.stringify(card);
    if (!cardStr || cardStr === this.lastSent) return;

    if (!this.ref) {
      // First send — create the card message
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
      // Update existing card via REST
      const success = await this.updateCardViaRest(card);
      if (success) {
        this.lastSent = cardStr;
      }
    }

    // Check if new content arrived while flushing
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
      // Check if we have content but no usage (stalled mid-stream)
      const hasContent = this.entries.some(
        (e) => (e.kind === "text" && e.text.length > 0) ||
               (e.kind === "tool-start" && e.children.length > 0),
      );
      const hasUsage = this.entries.some((e) => e.kind === "usage");
      if (hasContent && !hasUsage) {
        log.warn({ sessionId: this.sessionId }, "[SessionMessage] Stream stalled — adding cutoff notice");
        // Add a divider and notice as a text entry
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

  /** Finalize: clear stall timer, do a last flush. */
  async finalize(): Promise<MessageRef | null> {
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

  /** For split: finalize current message at limit, start fresh. */
  private split(): void {
    // Collect root text entries and check total length
    const rootText = this.entries
      .filter((e) => e.kind === "text")
      .map((e) => (e as { kind: "text"; text: string }).text)
      .join("");

    if (rootText.length <= MAX_ROOT_TEXT_LENGTH) return;

    // Find the split point
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
        // Non-text entries stay in the finalized message
        entriesToKeep.push(entry);
      }
    }

    log.info({ sessionId: this.sessionId, kept: entriesToKeep.length, overflow: entriesToOverflow.length }, "[SessionMessage] splitting");

    // Clear entries, add overflow text as new text entry
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

    // Finalize current message ref
    if (this.ref) {
      const card = this.buildCard();
      this.rateLimiter.enqueue(
        this.conversationId,
        () => this.updateCardViaRest(card).then(() => {}),
        `update:${this.ref.activityId}`,
      ).catch(() => {});
    }

    // Reset ref for new message
    this.ref = null;
    this.lastSent = "";
    this.requestFlush();
  }

  // ─── Legacy API (for adapter compat) ──────────────────────────────────────

  /** @deprecated — use addToolStart instead */
  setHeader(text: string): void {
    // No-op — header zone is gone. Tool calls use addToolStart.
  }

  /** @deprecated — use updateToolResult instead */
  setHeaderResult(text: string): void {
    // No-op — handled via updateToolResult
  }

  /** @deprecated — use addText instead */
  appendBody(text: string): void {
    this.addText(text);
  }

  /** @deprecated — use setUsage instead */
  setFooter(text: string): void {
    this.setUsage(text);
  }

  /** @deprecated — use appendFooter via setUsage instead */
  appendFooter(text: string): void {
    const current = this.getFooter();
    this.setUsage(current ? `${current} · ${text}` : text);
  }

  /** @deprecated — no-op, header is gone */
  clearHeader(): void {
    // No-op
  }
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

  /** Finalize and remove a session's message and plan ref. */
  async finalize(sessionId: string): Promise<MessageRef | null> {
    const msg = this.messages.get(sessionId);
    if (!msg) return null;
    this.messages.delete(sessionId);
    this.planRefs.delete(sessionId);
    return msg.finalize();
  }

  /** Get or set the plan message ref for a session. */
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