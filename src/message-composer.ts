/**
 * Message composer — manages a single Adaptive Card per session in Teams.
 *
 * Every message type renders into the same card, updated in-place via REST PUT.
 *
 * Entry types:
 *   - title:    Bold session name, always first
 *   - timed:    Two-level Container with live timer (tool, thinking)
 *   - info:     Two-level Container, no timer (error, system, mode, config, model)
 *   - text:     Root-level streamed text, always at bottom
 *   - plan:     Formatted plan list, updated in place
 *   - resource: Inline 📎 line (attachments, resources, resource links)
 *   - usage:    Italic footer, singleton
 *   - divider:  Horizontal rule
 *
 * Level 2 lines use a 3-column ColumnSet for true indentation:
 *   Column 1: spacer (20px)  |  Column 2: ⎿ (auto)  |  Column 3: content (stretch, wrap)
 */
import type { TurnContext } from "@microsoft/agents-hosting";
import { log } from "@openacp/plugin-sdk";
import { CardFactory } from "@microsoft/agents-hosting";
import type { ConversationRateLimiter } from "./rate-limiter.js";

// ─── Public types ─────────────────────────────────────────────────────────────

export interface MessageRef {
  activityId: string;
  conversationId: string;
  serviceUrl: string;
}

export type AcquireBotToken = () => Promise<string | null>;

// ─── Constants ────────────────────────────────────────────────────────────────

const MAX_ROOT_TEXT_LENGTH = 25_000;
const STALL_TIMEOUT = 120_000;

// ─── Entry Types ──────────────────────────────────────────────────────────────

type BodyEntry =
  | { id: string; kind: "title"; text: string }
  | { id: string; kind: "timed"; emoji: string; label: string; startedAt: number; result?: string; endedAt?: number; collapsible?: boolean }
  | { id: string; kind: "info"; emoji: string; label: string; detail: string }
  | { id: string; kind: "text"; text: string }
  | { id: string; kind: "plan"; entries: { content: string; status: string }[] }
  | { id: string; kind: "resource"; text: string }
  | { id: string; kind: "permission"; description: string; actions: { type: string; title: string; data: Record<string, unknown> }[]; resolved?: string }
  | { id: string; kind: "usage"; text: string }
  | { id: string; kind: "divider" };

// ─── ID generation ────────────────────────────────────────────────────────────

let _idCounter = 0;
function nextId(): string {
  return `e${++_idCounter}_${Date.now().toString(36)}`;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatElapsed(ms: number): string {
  if (ms < 1000) return `${ms}ms`;
  if (ms < 60_000) return `${(ms / 1000).toFixed(1)}s`;
  return `${Math.floor(ms / 60_000)}m ${Math.floor((ms % 60_000) / 1000)}s`;
}

export function escapeMd(text: string): string {
  // Adaptive Cards only support a subset of markdown in TextBlocks.
  // Backslash escapes render literally in Teams, so only escape brackets
  // (which trigger link syntax). Leave * and _ alone — they rarely form
  // valid bold/italic pairs in tool output and \* shows as literal \*.
  return text.replace(/[[\]]/g, "\\$&");
}

const PLAN_STATUS_ICONS: Record<string, string> = {
  completed: "✓",
  in_progress: "◼",
  pending: "◻",
};

// ─── Adaptive Card Builder ────────────────────────────────────────────────────

/**
 * Build a 2-column ColumnSet for level-1 headings.
 * Column 1: emoji (auto), Column 2: text (stretch, wrap).
 * Ensures long text wraps without breaking back to the left margin.
 */
export function buildLevel1(emoji: string, text: string): Record<string, unknown> {
  return {
    type: "ColumnSet",
    spacing: "None",
    columns: [
      {
        type: "Column",
        width: "auto",
        items: [{ type: "TextBlock", text: emoji, size: "Small", fontType: "Monospace", spacing: "None" }],
        verticalContentAlignment: "Top",
      },
      {
        type: "Column",
        width: "stretch",
        items: [{ type: "TextBlock", text, size: "Small", fontType: "Monospace", wrap: true, spacing: "None" }],
      },
    ],
  };
}

/**
 * Build a 3-column ColumnSet for indented level-2 content.
 * Column 1: spacer (20px), Column 2: ⎿ (auto), Column 3: content (stretch, wrap).
 */
export function buildLevel2(text: string, elapsed?: string, raw = false): Record<string, unknown> {
  const escaped = raw ? text : escapeMd(text);
  const content = elapsed ? `${escaped}  (${elapsed})` : escaped;
  return {
    type: "ColumnSet",
    spacing: "None",
    columns: [
      { type: "Column", width: "20px" },
      {
        type: "Column",
        width: "auto",
        items: [{
          type: "TextBlock",
          text: "⎿",
          size: "Small",
          fontType: "Monospace",
          spacing: "None",
        }],
        verticalContentAlignment: "Top",
      },
      {
        type: "Column",
        width: "stretch",
        items: [{
          type: "TextBlock",
          text: content,
          size: "Small",
          fontType: "Monospace",
          wrap: true,
          spacing: "None",
        }],
      },
    ],
  };
}

function buildCardBody(entries: BodyEntry[]): unknown[] {
  const blocks: unknown[] = [];
  let usageBlock: unknown | null = null;
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
          spacing: "Small",
        });
        break;

      case "timed": {
        const elapsed = entry.result
          ? formatElapsed((entry.endedAt ?? entry.startedAt) - entry.startedAt)
          : formatElapsed(now - entry.startedAt);

        if (entry.collapsible && entry.result) {
          // Completed thinking — collapsible with toggle
          const detailId = `detail-${entry.id}`;
          const showId = `show-${entry.id}`;
          const hideId = `hide-${entry.id}`;

          // "▶ Show" row — whole row is clickable, no separate button
          const showRow = {
            ...buildLevel1(entry.emoji, `${escapeMd(entry.label)}  (${elapsed})  ▶ Show`),
            id: showId,
            isVisible: true,
            selectAction: {
              type: "Action.ToggleVisibility",
              targetElements: [detailId, showId, hideId],
            },
          };

          // "▼ Hide" row — replaces show row when expanded
          const hideRow = {
            ...buildLevel1(entry.emoji, `${escapeMd(entry.label)}  (${elapsed})  ▼ Hide`),
            id: hideId,
            isVisible: false,
            selectAction: {
              type: "Action.ToggleVisibility",
              targetElements: [detailId, showId, hideId],
            },
          };

          blocks.push({
            type: "Container",
            spacing: "Small",
            items: [
              showRow,
              hideRow,
              // Collapsible detail — hidden by default
              { ...buildLevel2(entry.result, elapsed), id: detailId, isVisible: false },
            ],
          });
        } else {
          // Running (no result) or non-collapsible — standard rendering
          const headingText = entry.result
            ? escapeMd(entry.label)
            : `${escapeMd(entry.label)}…  (${elapsed})`;
          const items: unknown[] = [buildLevel1(entry.emoji, headingText)];
          if (entry.result) {
            items.push(buildLevel2(entry.result, elapsed));
          }
          blocks.push({ type: "Container", items, spacing: "Small" });
        }
        break;
      }

      case "info": {
        blocks.push({
          type: "Container",
          items: [
            buildLevel1(entry.emoji, escapeMd(entry.label)),
            buildLevel2(entry.detail),
          ],
          spacing: "Small",
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
          spacing: "Small",
        });
        break;

      case "plan": {
        const lines = entry.entries.map((e, i) => {
          const icon = PLAN_STATUS_ICONS[e.status] || "◻";
          const text = e.status === "completed" ? `~~${escapeMd(e.content)}~~` : escapeMd(e.content);
          return `${icon} ${i + 1}. ${text}`;
        });
        blocks.push({
          type: "TextBlock",
          text: `📋 Plan\n${lines.join("\n")}`,
          size: "Small",
          fontType: "Monospace",
          wrap: true,
          spacing: "Small",
        });
        break;
      }

      case "resource":
        blocks.push({
          type: "TextBlock",
          text: entry.text,
          size: "Small",
          fontType: "Monospace",
          wrap: true,
          spacing: "Small",
        });
        break;

      case "permission": {
        const permItems: unknown[] = [];
        if (entry.resolved) {
          const icon = entry.resolved.startsWith("Denied") ? "❌" : "✅";
          permItems.push(buildLevel1(icon, `Permission — ${escapeMd(entry.resolved)}`));
          permItems.push(buildLevel2(entry.description));
        } else {
          permItems.push(buildLevel1("🔐", "Permission"));
          permItems.push(buildLevel2(entry.description));
          permItems.push({
            type: "ActionSet",
            spacing: "Small",
            actions: entry.actions,
          });
        }
        blocks.push({ type: "Container", items: permItems, spacing: "Small" });
        break;
      }

      case "usage":
        usageBlock = {
          type: "TextBlock",
          text: `*${escapeMd(entry.text)}*`,
          isSubtle: true,
          size: "Small",
          fontType: "Monospace",
          spacing: "Small",
        };
        break;

      case "divider":
        blocks.push({
          type: "TextBlock",
          text: "─".repeat(30),
          size: "Small",
          fontType: "Monospace",
          spacing: "Small",
        });
        break;
    }
  }

  if (usageBlock) blocks.push(usageBlock);
  return blocks;
}

// ─── SessionMessage ───────────────────────────────────────────────────────────

/** Dot animation frames: Working. → Working.. → Working... → Working.. → repeat */
const WORKING_FRAMES = ["Working.", "Working..", "Working...", "Working.."];

export class SessionMessage {
  private entries: BodyEntry[] = [];
  private titleId: string | null = null;
  private usageId: string | null = null;
  private planId: string | null = null;
  private thinkingActive: string | null = null;
  private thinkingText = "";
  private working = true;
  private workingFrame = 0;
  private destroyed = false;
  private ref: MessageRef | null = null;
  private creating = false;
  private sealed = false;
  private lastSent = "";
  private stallTimer?: ReturnType<typeof setTimeout>;
  private tickInterval?: ReturnType<typeof setInterval>;
  private emptyCardTimer?: ReturnType<typeof setTimeout>;

  constructor(
    private context: TurnContext,
    private conversationId: string,
    private sessionId: string,
    private rateLimiter: ConversationRateLimiter,
    private acquireBotToken: AcquireBotToken,
  ) {
    // Start the working animation immediately — enqueue flush with no debounce
    this.usageId = nextId();
    this.entries.push({ id: this.usageId, kind: "usage", text: WORKING_FRAMES[0] });
    this.startTickInterval();
    this.rateLimiter.enqueue(
      this.conversationId,
      () => this.flush(),
      `new:${this.sessionId}`,
    ).catch(() => {});

    // If no real content arrives within 30s, delete the empty card
    this.emptyCardTimer = setTimeout(() => this.deleteIfEmpty(), 30_000);
    if (this.emptyCardTimer.unref) this.emptyCardTimer.unref();
  }

  get isSealed(): boolean {
    return this.sealed;
  }

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

  // ─── Empty card cleanup ──────────────────────────────────────────────────

  /** Cancel the empty-card timer (called when real content arrives). */
  private cancelEmptyCardTimer(): void {
    if (this.emptyCardTimer) {
      clearTimeout(this.emptyCardTimer);
      this.emptyCardTimer = undefined;
    }
  }

  /** If no real content was added, delete the card activity and clean up. */
  private async deleteIfEmpty(): Promise<void> {
    this.emptyCardTimer = undefined;
    if (this.entries.some((e) => e.kind !== "usage")) return;
    this.destroyed = true;

    // Delete the card activity if it was sent
    if (this.ref) {
      const token = await this.acquireBotToken();
      // Re-check after await — content may have arrived while token was fetched
      if (this.entries.some((e) => e.kind !== "usage")) {
        this.destroyed = false;
        return;
      }
      if (token) {
        const url = `${this.ref.serviceUrl}/v3/conversations/${encodeURIComponent(this.ref.conversationId)}/activities/${encodeURIComponent(this.ref.activityId)}`;
        try {
          await fetch(url, { method: "DELETE", headers: { "Authorization": `Bearer ${token}` } });
        } catch { /* best effort */ }
      }
      this.ref = null;
    }

    this.stopTickInterval();
    this.working = false;
    this.entries = [];
    log.debug({ sessionId: this.sessionId }, "[SessionMessage] Empty card deleted after timeout");
  }

  // ─── Entry API ──────────────────────────────────────────────────────────

  /** Set the persistent session title (bold, always first entry). */
  setTitle(text: string): void {
    this.cancelEmptyCardTimer();
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
   * Start a timed entry (tool or thinking). Shows emoji + label with live elapsed timer.
   * Returns the entry ID for pairing with addTimedResult().
   */
  addTimedStart(emoji: string, label: string): string {
    this.cancelEmptyCardTimer();
    const id = nextId();
    this.entries.push({ id, kind: "timed", emoji, label, startedAt: Date.now() });
    this.resetStallTimer();
    this.startTickInterval();
    this.requestFlush();
    return id;
  }

  /**
   * Close a timed entry by setting its result. Stops the tick interval if no
   * other timed entries are still running.
   */
  addTimedResult(id: string, result: string): void {
    const entry = this.findEntry(id);
    if (entry && entry.kind === "timed") {
      entry.result = result;
      entry.endedAt = Date.now();
    } else {
      // Orphan result — create a standalone completed entry
      this.entries.push({ id: nextId(), kind: "timed", emoji: "🔧", label: result, startedAt: Date.now(), result, endedAt: Date.now() });
    }
    const hasRunning = this.entries.some((e) => e.kind === "timed" && !e.result);
    if (!hasRunning && !this.working) this.stopTickInterval();
    this.requestFlush();
  }

  /**
   * Start or accumulate thinking text. The first call creates a timed entry;
   * subsequent calls append to the pending result text. Call closeActiveThinking()
   * to finalize.
   */
  addThinking(text: string): void {
    this.cancelEmptyCardTimer();
    if (!this.thinkingActive) {
      const id = this.addTimedStart("☁️", "Thinking");
      // Mark as collapsible so the card builder adds a toggle after completion
      const entry = this.findEntry(id);
      if (entry && entry.kind === "timed") entry.collapsible = true;
      this.thinkingActive = id;
      this.thinkingText = text;
    } else {
      this.thinkingText += ` ${text}`;
    }
    this.requestFlush();
  }

  /**
   * Close the active thinking entry (if any). Called by non-thought handlers
   * so thinking ends when the next event type arrives.
   */
  closeActiveThinking(): void {
    if (!this.thinkingActive) return;
    const text = this.thinkingText.trim() || "…";
    this.addTimedResult(this.thinkingActive, text);
    this.thinkingActive = null;
    this.thinkingText = "";
  }

  /** Add a one-shot info entry (error, system, mode, config, model). */
  addInfo(emoji: string, label: string, detail: string): void {
    this.cancelEmptyCardTimer();
    this.entries.push({ id: nextId(), kind: "info", emoji, label, detail });
    this.requestFlush();
  }

  /**
   * Add a permission request entry with action buttons. Returns the entry ID
   * for resolving later with resolvePermission().
   */
  addPermission(description: string, actions: { type: string; title: string; data: Record<string, unknown> }[]): string {
    this.cancelEmptyCardTimer();
    const id = nextId();
    this.entries.push({ id, kind: "permission", description, actions });
    this.requestFlush();
    return id;
  }

  /** Resolve a permission entry — replace buttons with the result text. */
  resolvePermission(id: string, result: string): void {
    const entry = this.findEntry(id);
    if (entry && entry.kind === "permission") {
      entry.resolved = result;
      entry.actions = [];
      this.requestFlush();
    }
  }

  /** Add or replace the plan entry. */
  setPlan(entries: { content: string; status: string }[]): void {
    this.cancelEmptyCardTimer();
    if (this.planId) {
      const entry = this.findEntry(this.planId);
      if (entry && entry.kind === "plan") {
        entry.entries = entries;
        this.requestFlush();
        return;
      }
    }
    this.planId = nextId();
    this.entries.push({ id: this.planId, kind: "plan", entries });
    this.requestFlush();
  }

  /** Add text — always appended at root level. */
  addText(text: string): void {
    if (!text) return;
    this.cancelEmptyCardTimer();
    const lastText = this.entries.filter((e) => e.kind === "text").at(-1);
    if (lastText && lastText.kind === "text") {
      lastText.text += text;
    } else {
      this.entries.push({ id: nextId(), kind: "text", text });
    }
    this.resetStallTimer();
    this.requestFlush();
    if (this.ref) this.checkSplit();
  }

  /** Add an inline resource line (📎 prefix). */
  addResource(text: string): void {
    this.cancelEmptyCardTimer();
    this.entries.push({ id: nextId(), kind: "resource", text });
    this.requestFlush();
  }

  /** Set or replace the usage footer. Stops the working animation. */
  setUsage(text: string): void {
    this.working = false;
    if (this.usageId) {
      const entry = this.findEntry(this.usageId);
      if (entry && entry.kind === "usage") entry.text = text;
    } else {
      this.usageId = nextId();
      this.entries.push({ id: this.usageId, kind: "usage", text });
    }
    const hasRunning = this.entries.some((e) => e.kind === "timed" && !e.result);
    if (!hasRunning) this.stopTickInterval();
    this.requestFlush();
  }

  /** Add a divider entry. */
  appendDivider(): void {
    this.entries.push({ id: nextId(), kind: "divider" });
    this.requestFlush();
  }

  // ─── Entry helpers ──────────────────────────────────────────────────────

  private findEntry(id: string): BodyEntry | undefined {
    return this.entries.find((e) => e.id === id);
  }

  // ─── Periodic tick for elapsed time updates ─────────────────────────────

  private startTickInterval(): void {
    if (this.tickInterval) return;
    this.tickInterval = setInterval(() => {
      // Advance working dots animation
      if (this.working && this.usageId) {
        this.workingFrame = (this.workingFrame + 1) % WORKING_FRAMES.length;
        const entry = this.findEntry(this.usageId);
        if (entry && entry.kind === "usage") entry.text = WORKING_FRAMES[this.workingFrame];
      }

      const hasRunning = this.entries.some((e) => e.kind === "timed" && !e.result);
      if (!hasRunning && !this.working) {
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

  // ─── Card building ──────────────────────────────────────────────────────

  private buildCard(): Record<string, unknown> {
    const body = buildCardBody(this.entries);
    return {
      type: "AdaptiveCard",
      version: "1.4",
      body: body.length > 0 ? body : [{ type: "TextBlock", text: "…" }],
      msteams: { width: "Full" },
    };
  }

  // ─── Flush / Rate limiting ──────────────────────────────────────────────

  private flushTimer?: ReturnType<typeof setTimeout>;

  private requestFlush(): void {
    if (this.flushTimer) clearTimeout(this.flushTimer);
    this.flushTimer = setTimeout(() => {
      this.flushTimer = undefined;
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
    if (this.destroyed) return;
    const card = this.buildCard();
    const cardStr = JSON.stringify(card);
    if (!cardStr || cardStr === this.lastSent) return;

    if (!this.ref) {
      // Guard against concurrent creates (flush + finalize racing)
      if (this.creating) return;
      this.creating = true;
      try {
        const result = await sendCard(this.context, card) as { id?: string } | undefined;
        if (result?.id) {
          this.ref = {
            activityId: result.id,
            conversationId: this.context.activity.conversation?.id as string,
            serviceUrl: this.context.activity.serviceUrl as string,
          };
        }
        this.lastSent = cardStr;
      } finally {
        this.creating = false;
      }
    } else {
      const success = await this.updateCardViaRest(card);
      if (success) this.lastSent = cardStr;
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
        headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
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

  // ─── Stall timer ────────────────────────────────────────────────────────

  private resetStallTimer(): void {
    if (this.stallTimer) clearTimeout(this.stallTimer);
    this.stallTimer = setTimeout(() => {
      const hasContent = this.entries.some((e) => e.kind === "text" && e.text.length > 0);
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

  // ─── Finalize ───────────────────────────────────────────────────────────

  async finalize(): Promise<MessageRef | null> {
    this.cancelEmptyCardTimer();
    this.closeActiveThinking();
    this.working = false;
    this.sealed = true;
    this.stopTickInterval();
    if (this.flushTimer) {
      clearTimeout(this.flushTimer);
      this.flushTimer = undefined;
    }
    if (this.stallTimer) {
      clearTimeout(this.stallTimer);
      this.stallTimer = undefined;
    }

    // Wait for any in-flight create to finish before finalizing
    if (this.creating) {
      await new Promise<void>((r) => {
        const check = setInterval(() => {
          if (!this.creating) { clearInterval(check); r(); }
        }, 50);
      });
    }

    const card = this.buildCard();
    const cardStr = JSON.stringify(card);
    if (cardStr && cardStr !== this.lastSent) {
      if (!this.ref) {
        this.creating = true;
        try {
          const result = await sendCard(this.context, card) as { id?: string } | undefined;
          if (result?.id) {
            this.ref = {
              activityId: result.id,
              conversationId: this.context.activity.conversation?.id as string,
              serviceUrl: this.context.activity.serviceUrl as string,
            };
          }
        } finally {
          this.creating = false;
        }
      } else {
        await this.updateCardViaRest(card);
      }
      this.lastSent = cardStr;
    }

    return this.ref;
  }

  /** Strip a pattern from root text entries. */
  async stripPattern(pattern: RegExp): Promise<void> {
    for (const entry of this.entries) {
      if (entry.kind === "text") {
        try {
          entry.text = entry.text.replace(pattern, "").trim();
        } catch { /* leave unchanged */ }
      }
    }
  }

  // ─── Text overflow / split ──────────────────────────────────────────────

  private checkSplit(): void {
    const rootText = this.entries
      .filter((e) => e.kind === "text")
      .map((e) => (e as { kind: "text"; text: string }).text)
      .join("");
    if (rootText.length <= MAX_ROOT_TEXT_LENGTH) return;
    this.split();
  }

  private split(): void {
    const rootText = this.entries
      .filter((e) => e.kind === "text")
      .map((e) => (e as { kind: "text"; text: string }).text)
      .join("");
    if (rootText.length <= MAX_ROOT_TEXT_LENGTH) return;

    let accLen = 0;
    const keep: BodyEntry[] = [];
    const overflow: BodyEntry[] = [];

    for (const entry of this.entries) {
      if (entry.kind === "text") {
        if (accLen + entry.text.length <= MAX_ROOT_TEXT_LENGTH) {
          keep.push(entry);
          accLen += entry.text.length;
        } else {
          overflow.push(entry);
        }
      } else {
        keep.push(entry);
      }
    }

    log.info({ sessionId: this.sessionId, kept: keep.length, overflow: overflow.length }, "[SessionMessage] splitting");

    const overflowText = overflow
      .filter((e) => e.kind === "text")
      .map((e) => (e as { kind: "text"; text: string }).text)
      .join("");

    // Update the old card with only the kept entries (no overflow).
    // Capture the old ref before nulling it — the rate limiter closure
    // must use the old ref, not whatever this.ref points to later.
    this.entries = keep;
    const oldRef = this.ref;
    if (oldRef) {
      const card = this.buildCard();
      this.rateLimiter.enqueue(
        this.conversationId,
        async () => {
          const token = await this.acquireBotToken();
          if (!token) return;
          const url = `${oldRef.serviceUrl}/v3/conversations/${encodeURIComponent(oldRef.conversationId)}/activities/${encodeURIComponent(oldRef.activityId)}`;
          try {
            await fetch(url, {
              method: "PUT",
              headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
              body: JSON.stringify({ type: "message", attachments: [CardFactory.adaptiveCard(card)] }),
            });
          } catch { /* best effort */ }
        },
        `update:${oldRef.activityId}`,
      ).catch(() => {});
    }

    // Start a new card with only the overflow text — reset all entry ID references
    // so subsequent calls create new entries in the new card
    this.ref = null;
    this.lastSent = "";
    this.titleId = null;
    this.usageId = null;
    this.planId = null;
    this.entries = overflowText
      ? [{ id: nextId(), kind: "text" as const, text: overflowText }]
      : [];
    this.requestFlush();
  }
}

// ─── SessionMessageManager ────────────────────────────────────────────────────

export class SessionMessageManager {
  private messages = new Map<string, SessionMessage>();

  constructor(
    private rateLimiter: ConversationRateLimiter,
    private acquireBotToken: AcquireBotToken,
  ) {}

  getOrCreate(sessionId: string, context: TurnContext): SessionMessage {
    let msg = this.messages.get(sessionId);
    if (msg && msg.isSealed) {
      // Previous turn's card was soft-finalized — remove it and start fresh
      this.messages.delete(sessionId);
      msg = undefined;
    }
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
    return msg.finalize();
  }

  /** Flush the final card state but keep the SessionMessage in the map for late events. */
  async softFinalize(sessionId: string): Promise<MessageRef | null> {
    const msg = this.messages.get(sessionId);
    if (!msg) return null;
    return msg.finalize();
  }

  cleanup(sessionId: string): void {
    const msg = this.messages.get(sessionId);
    if (msg) msg.finalize().catch(() => {});
    this.messages.delete(sessionId);
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

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
