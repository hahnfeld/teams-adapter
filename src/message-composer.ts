/**
 * Message composer — manages a single "main message" per session in Teams.
 *
 * The main message has four zones:
 *   TITLE   (persistent, bold) — session name, survives finalize
 *   HEADER  (ephemeral, italic) — tool_call, tool_update, thought
 *   BODY    (persistent, streamed) — text, attachments
 *   FOOTER  (persistent, italic) — usage, session_end
 *
 * All zones are composed into a single markdown string and sent/updated
 * as one Teams message via the rate limiter.
 */
import type { TurnContext } from "@microsoft/agents-hosting";
import { log } from "@openacp/plugin-sdk";
import { splitMessage } from "./formatting.js";
import { sendText } from "./send-utils.js";
import type { ConversationRateLimiter } from "./rate-limiter.js";

const MAX_BODY_LENGTH = 25_000;
/** How long to wait with no activity (body text or header changes) before warning about truncation. */
const STALL_TIMEOUT = 120_000;

export interface MessageRef {
  activityId: string;
  conversationId: string;
  serviceUrl: string;
}

export type AcquireBotToken = () => Promise<string | null>;

/** Escape asterisks to prevent breaking markdown italic/bold spans. */
function escapeEmphasis(text: string): string {
  return text.replace(/\*/g, "\\*");
}

/**
 * Normalize newlines for Teams rendering.
 * Teams collapses single \n in markdown — use \n\n for line breaks.
 */
function teamsNewlines(text: string): string {
  return text.replace(/(?<!\n)\n(?!\n)/g, "\n\n");
}

/**
 * A single main message with title/header/body/footer zones.
 */
export class SessionMessage {
  private title: string | null = null;
  private header: string | null = null;
  private body = "";
  private footer: string | null = null;
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
    return this.body;
  }

  getFooter(): string | null {
    return this.footer;
  }

  /** Set the persistent session title (bold, survives finalize). */
  setTitle(text: string): void {
    this.title = text;
    this.requestFlush();
  }

  /** Max header length — tool notifications can include file paths and summaries. */
  private static readonly MAX_HEADER_LENGTH = 300;

  /** Replace the ephemeral header (tool_call, thought, etc.). Reset stall timer since tool activity counts as activity. */
  setHeader(text: string): void {
    // Truncate to first line, then cap length — headers are status indicators, not content
    const firstLine = text.split("\n")[0];
    this.header = firstLine.length > SessionMessage.MAX_HEADER_LENGTH
      ? firstLine.slice(0, SessionMessage.MAX_HEADER_LENGTH) + "..."
      : firstLine;
    this.resetStallTimer();
    this.requestFlush();
  }

  /** Clear the ephemeral header and flush the update. */
  clearHeader(): void {
    if (this.header === null) return;
    this.header = null;
    this.requestFlush();
  }

  /** Append text to the body (streaming text chunks). */
  appendBody(text: string): void {
    if (!text) return;
    this.body += text;
    this.resetStallTimer();

    // Check if body needs splitting
    if (this.body.length > MAX_BODY_LENGTH && this.ref) {
      this.split();
      return;
    }

    this.requestFlush();
  }

  /** Set the persistent footer (usage, completion). */
  setFooter(text: string): void {
    this.footer = text;
    this.requestFlush();
  }

  /** Append to the existing footer (e.g., adding "Task completed" after usage). */
  appendFooter(text: string): void {
    this.footer = this.footer ? `${this.footer} · ${text}` : text;
    this.requestFlush();
  }

  /**
   * Close any unclosed code fences in the body.
   * If the body has an odd number of ``` markers, the footer would
   * be swallowed into the code block — append a closing fence.
   */
  private static closeCodeFences(text: string): string {
    const fenceCount = (text.match(/^```/gm) || []).length;
    if (fenceCount % 2 !== 0) {
      return text + "\n```";
    }
    return text;
  }

  /** Compose the four zones into a single markdown string. */
  compose(): string {
    const parts: string[] = [];

    if (this.title) {
      parts.push(`**${escapeEmphasis(this.title)}**`);
    }

    if (this.header) {
      parts.push(`*${escapeEmphasis(this.header)}*`);
    }

    if (this.title || this.header) {
      parts.push("---");
    }

    if (this.body) {
      parts.push(SessionMessage.closeCodeFences(this.body));
    }

    if (this.footer) {
      if (this.body) parts.push("---");
      parts.push(`*${escapeEmphasis(this.footer)}*`);
    }

    return parts.join("\n\n");
  }

  /** Request a flush through the rate limiter. */
  requestFlush(): void {
    const composed = this.compose();
    if (!composed) return;
    if (composed === this.lastSent) return;

    // Coalescing key: activityId for updates, session for new sends
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
    const composed = this.compose();
    if (!composed || composed === this.lastSent) return;

    if (!this.ref) {
      // First send — create the message
      const result = await sendText(this.context, composed) as { id?: string } | undefined;
      if (result?.id) {
        this.ref = {
          activityId: result.id,
          conversationId: this.context.activity.conversation?.id as string,
          serviceUrl: this.context.activity.serviceUrl as string,
        };
      }
      this.lastSent = composed;
    } else {
      // Update existing message via REST
      const success = await this.updateViaRest(composed);
      if (success) {
        this.lastSent = composed;
      }
    }

    // Check if state changed during the flush (new chunks arrived while we were awaiting)
    const current = this.compose();
    if (current && current !== this.lastSent) {
      this.requestFlush();
    }
  }

  private async updateViaRest(text: string): Promise<boolean> {
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
          text: teamsNewlines(text),
          textFormat: "markdown",
        }),
      });

      if (!response.ok) {
        log.warn({ status: response.status, sessionId: this.sessionId }, "[SessionMessage] REST update failed");
        return false;
      }
      return true;
    } catch (err) {
      log.warn({ err, sessionId: this.sessionId }, "[SessionMessage] REST update error");
      return false;
    }
  }

  /** Split: finalize current message at body limit, start fresh. */
  private split(): void {
    const finalBody = this.body.slice(0, MAX_BODY_LENGTH);
    const overflow = this.body.slice(MAX_BODY_LENGTH);

    log.info({ sessionId: this.sessionId, finalLen: finalBody.length, overflowLen: overflow.length }, "[SessionMessage] splitting");

    // Clear ephemeral header on the finalized message; footer stays per spec
    this.header = null;
    this.body = finalBody;

    // Final flush of the current message (with footer preserved)
    const composed = this.compose();
    if (this.ref && composed !== this.lastSent) {
      this.rateLimiter.enqueue(
        this.conversationId,
        () => this.updateViaRest(composed).then(() => {}),
        `update:${this.ref.activityId}`,
      ).catch(() => {});
    }

    // Reset for a new message — footer carries over to new message
    this.ref = null;
    this.lastSent = "";
    this.body = overflow;

    if (overflow) {
      this.requestFlush();
    }
  }

  private resetStallTimer(): void {
    if (this.stallTimer) clearTimeout(this.stallTimer);
    this.stallTimer = setTimeout(() => {
      if (this.body && !this.footer) {
        log.warn({ sessionId: this.sessionId }, "[SessionMessage] Stream stalled — adding cutoff notice");
        this.header = null;
        // Use appendBody so split check is applied
        this.appendBody("\n\n---\n_Response was cut short — the model likely reached its output token limit. Send a follow-up message to continue._");
      }
    }, STALL_TIMEOUT);
    if (this.stallTimer.unref) this.stallTimer.unref();
  }

  /** Finalize: clear stall timer, clear ephemeral header, do a last flush. */
  async finalize(): Promise<MessageRef | null> {
    if (this.stallTimer) {
      clearTimeout(this.stallTimer);
      this.stallTimer = undefined;
    }

    // Clear ephemeral header on finalize; title and footer persist
    this.header = null;

    // Final flush
    const composed = this.compose();
    if (composed && composed !== this.lastSent) {
      if (!this.ref) {
        const result = await sendText(this.context, composed) as { id?: string } | undefined;
        if (result?.id) {
          this.ref = {
            activityId: result.id,
            conversationId: this.context.activity.conversation?.id as string,
            serviceUrl: this.context.activity.serviceUrl as string,
          };
        }
      } else {
        await this.updateViaRest(composed);
      }
      this.lastSent = composed;
    }

    return this.ref;
  }

  async stripPattern(pattern: RegExp): Promise<void> {
    if (!this.body) return;
    try {
      this.body = this.body.replace(pattern, "").trim();
    } catch { /* leave unchanged */ }
  }
}

/**
 * Manages SessionMessage instances and plan refs across sessions.
 */
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
