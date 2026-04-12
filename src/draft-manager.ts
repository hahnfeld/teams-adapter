import type { TurnContext } from "@microsoft/agents-hosting";
import { log } from "@openacp/plugin-sdk";
import { splitMessage } from "./formatting.js";
import { sendText } from "./send-utils.js";

/** First flush fires quickly so the user sees content fast */
const FIRST_FLUSH_INTERVAL = 500;
/** Subsequent flushes are slower to avoid rate limits (Teams: 7 ops/sec/conversation) */
const UPDATE_FLUSH_INTERVAL = 2000;
const MAX_DISPLAY_LENGTH = 25000;
/** If no new text arrives for this long, consider the stream stalled */
const STALL_TIMEOUT = 30_000;

export interface MessageRef {
  activityId?: string;
  conversationId?: string;
  serviceUrl?: string;
}

/** Function signature for acquiring a bot token — injected from the adapter */
export type AcquireBotToken = () => Promise<string | null>;

export class TeamsMessageDraft {
  private buffer: string = "";
  private ref?: MessageRef;
  private flushTimer?: ReturnType<typeof setTimeout>;
  private flushPromise: Promise<void> = Promise.resolve();
  private lastSentBuffer: string = "";
  private firstFlushPending = false;
  private finalizing = false;
  private lastAppendTime = 0;
  private stallTimer?: ReturnType<typeof setTimeout>;
  /** After finalize, holds the last activity ref + text for post-finalize edits */
  private _finalRef?: { activityId: string; conversationId: string; serviceUrl: string; text: string };

  constructor(
    private context: TurnContext,
    private sendQueue: { enqueue<T>(fn: () => Promise<T>, opts?: { type?: string }): Promise<T | undefined> },
    private sessionId: string,
    private acquireBotToken?: AcquireBotToken,
  ) {}

  /** Update the TurnContext to the latest inbound turn (prevents stale context refs) */
  updateContext(context: TurnContext): void {
    this.context = context;
  }

  append(text: string): void {
    if (!text) return;
    this.buffer += text;
    this.lastAppendTime = Date.now();
    this.scheduleFlush();
    this.resetStallTimer();
  }

  private resetStallTimer(): void {
    if (this.stallTimer) clearTimeout(this.stallTimer);
    this.stallTimer = setTimeout(() => {
      if (this.buffer && !this.finalizing) {
        log.warn({ sessionId: this.sessionId, bufLen: this.buffer.length }, "[TeamsMessageDraft] Stream stalled — finalizing");
        this.buffer += "\n\n---\n_Response was cut short — the model likely reached its output token limit. Send a follow-up message to continue._";
        this.finalize().catch(() => {});
      }
    }, STALL_TIMEOUT);
    if (this.stallTimer.unref) this.stallTimer.unref();
  }

  getBuffer(): string {
    return this.buffer;
  }

  /** Return the last finalized activity ref for post-finalize edits (e.g., appending "Task completed") */
  getFinalRef(): { activityId: string; conversationId: string; serviceUrl: string; text: string } | undefined {
    return this._finalRef;
  }

  private scheduleFlush(): void {
    if (this.flushTimer) return;
    const interval = this.ref?.activityId ? UPDATE_FLUSH_INTERVAL : FIRST_FLUSH_INTERVAL;
    this.flushTimer = setTimeout(() => {
      this.flushTimer = undefined;
      this.flushPromise = this.flushPromise.then(() => this.flush()).catch(() => {});
    }, interval);
  }

  async flush(): Promise<void> {
    if (!this.buffer) return;
    if (this.firstFlushPending) return;
    log.debug({ sessionId: this.sessionId, bufLen: this.buffer.length, hasRef: !!this.ref?.activityId }, "[TeamsMessageDraft] flush");

    // If buffer exceeds limit, finalize the current message with its portion
    // and start a new streaming message for the overflow.
    if (this.buffer.length > MAX_DISPLAY_LENGTH && this.ref?.activityId) {
      const finalChunk = this.buffer.slice(0, MAX_DISPLAY_LENGTH);
      const overflow = this.buffer.slice(MAX_DISPLAY_LENGTH);

      // Update current message with its final content
      log.info({ sessionId: this.sessionId, finalChunkLen: finalChunk.length, overflowLen: overflow.length }, "[TeamsMessageDraft] splitting to new message");
      await this.updateActivityViaRest(finalChunk).catch((err) => {
        log.warn({ err, sessionId: this.sessionId }, "[TeamsMessageDraft] split: update failed");
      });

      // Reset for a new message
      this.ref = undefined;
      this.lastSentBuffer = "";
      this.firstFlushPending = false;
      this.buffer = overflow;

      // Immediately send the overflow as a new message to resume streaming
      if (overflow) {
        this.firstFlushPending = true;
        try {
          const result = await this.sendQueue.enqueue(
            () => sendText(this.context, overflow) as Promise<unknown>,
            { type: "other" },
          );
          if (result) {
            this.ref = {
              activityId: (result as { id?: string }).id,
              conversationId: this.context.activity.conversation?.id as string | undefined,
              serviceUrl: this.context.activity.serviceUrl as string | undefined,
            };
            this.lastSentBuffer = overflow;
          }
        } catch (err) {
          log.warn({ err, sessionId: this.sessionId }, "[TeamsMessageDraft] flush: overflow send failed");
        } finally {
          this.firstFlushPending = false;
        }
      }
      return;
    }

    const snapshot = this.buffer;

    if (!this.ref?.activityId) {
      // First message — send via context.send() to get the activityId
      this.firstFlushPending = true;
      try {
        const result = await this.sendQueue.enqueue(
          () => sendText(this.context, snapshot) as Promise<unknown>,
          { type: "other" },
        );
        if (result) {
          this.ref = {
            activityId: (result as { id?: string }).id,
            conversationId: this.context.activity.conversation?.id as string | undefined,
            serviceUrl: this.context.activity.serviceUrl as string | undefined,
          };
          this.lastSentBuffer = snapshot;
        }
      } catch (err) {
        log.warn({ err, sessionId: this.sessionId }, "[TeamsMessageDraft] flush: initial send failed");
      } finally {
        this.firstFlushPending = false;
      }
    } else {
      // Subsequent updates — use Bot Framework REST API to edit the message
      if (snapshot === this.lastSentBuffer) return;

      try {
        const success = await this.updateActivityViaRest(snapshot);
        if (success) {
          this.lastSentBuffer = snapshot;
        } else {
          log.warn({ sessionId: this.sessionId, bufLen: snapshot.length }, "[TeamsMessageDraft] flush: REST update returned false");
        }
      } catch (err) {
        log.warn({ err, sessionId: this.sessionId }, "[TeamsMessageDraft] flush: update failed");
      }
    }
  }

  /**
   * Update an existing message via the Bot Framework REST API.
   * Bypasses the teams.apps SDK context which doesn't support updateActivity.
   */
  private async updateActivityViaRest(text: string): Promise<boolean> {
    if (!this.ref?.activityId || !this.ref.serviceUrl || !this.ref.conversationId) return false;
    if (!this.acquireBotToken) return false;

    const token = await this.acquireBotToken();
    if (!token) return false;

    const url = `${this.ref.serviceUrl}/v3/conversations/${encodeURIComponent(this.ref.conversationId)}/activities/${encodeURIComponent(this.ref.activityId)}`;

    const response = await fetch(url, {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`,
      },
      body: JSON.stringify({
        type: "message",
        text: text.replace(/(?<!\n)\n(?!\n)/g, "\n\n"), // Teams newline normalization
        textFormat: "markdown",
      }),
    });

    if (!response.ok) {
      log.warn({ status: response.status, sessionId: this.sessionId }, "[TeamsMessageDraft] REST updateActivity failed");
      return false;
    }

    return true;
  }

  private _saveFinalRef(text: string): void {
    if (this.ref?.activityId && this.ref.conversationId && this.ref.serviceUrl) {
      this._finalRef = {
        activityId: this.ref.activityId,
        conversationId: this.ref.conversationId,
        serviceUrl: this.ref.serviceUrl,
        text,
      };
    }
  }

  async stripPattern(pattern: RegExp): Promise<void> {
    if (!this.buffer) return;
    try {
      this.buffer = this.buffer.replace(pattern, "").trim();
    } catch {
      // Regex failed — leave buffer unchanged
    }
  }

  async finalize(): Promise<void> {
    if (this.finalizing) return;
    this.finalizing = true;
    if (this.stallTimer) { clearTimeout(this.stallTimer); this.stallTimer = undefined; }
    try {
      await this._finalizeInner();
    } finally {
      this.finalizing = false;
    }
  }

  private async _finalizeInner(): Promise<void> {
    if (this.flushTimer) {
      clearTimeout(this.flushTimer);
      this.flushTimer = undefined;
    }

    // Wait for any in-flight flush to complete
    await this.flushPromise;

    if (!this.buffer) return;

    // If we have an activityId and the buffer fits in one message, do a final update
    if (this.ref?.activityId && this.buffer.length <= MAX_DISPLAY_LENGTH) {
      if (this.buffer !== this.lastSentBuffer) {
        try {
          const success = await this.updateActivityViaRest(this.buffer);
          if (success) {
            this._saveFinalRef(this.buffer);
            return;
          }
        } catch {
          // Fall through to send as new message
        }
      } else {
        this._saveFinalRef(this.buffer);
        return; // Already sent, nothing to update
      }
    }

    // Buffer exceeds single message limit or update failed — send as new message(s).
    // If we have an existing streaming message, update it with the first chunk,
    // then send the rest as new messages.
    const chunks = splitMessage(this.buffer, MAX_DISPLAY_LENGTH);

    let lastChunkText = "";
    for (let i = 0; i < chunks.length; i++) {
      const content = chunks[i];
      try {
        if (i === 0 && this.ref?.activityId) {
          const success = await this.updateActivityViaRest(content);
          if (!success) {
            const result = await this.sendQueue.enqueue(
              () => sendText(this.context, content) as Promise<unknown>,
              { type: "other" },
            );
            if (result) {
              this.ref = {
                activityId: (result as { id?: string }).id,
                conversationId: this.context.activity.conversation?.id as string | undefined,
                serviceUrl: this.context.activity.serviceUrl as string | undefined,
              };
            }
          }
        } else {
          const result = await this.sendQueue.enqueue(
            () => sendText(this.context, content) as Promise<unknown>,
            { type: "other" },
          );
          if (result) {
            this.ref = {
              activityId: (result as { id?: string }).id,
              conversationId: this.context.activity.conversation?.id as string | undefined,
              serviceUrl: this.context.activity.serviceUrl as string | undefined,
            };
          }
        }
        lastChunkText = content;
      } catch (err) {
        log.warn({ err, sessionId: this.sessionId, chunk: i }, "[TeamsMessageDraft] finalize: chunk send failed");
      }
    }
    if (lastChunkText) this._saveFinalRef(lastChunkText);
  }
}

export class TeamsDraftManager {
  private drafts = new Map<string, TeamsMessageDraft>();

  constructor(
    private sendQueue: { enqueue<T>(fn: () => Promise<T>, opts?: { type?: string }): Promise<T | undefined> },
    private acquireBotToken?: AcquireBotToken,
  ) {}

  getOrCreate(sessionId: string, context: TurnContext): TeamsMessageDraft {
    let draft = this.drafts.get(sessionId);
    if (!draft) {
      draft = new TeamsMessageDraft(context, this.sendQueue, sessionId, this.acquireBotToken);
      this.drafts.set(sessionId, draft);
    } else {
      // Re-bind to the latest TurnContext to prevent stale context references
      // when a new inbound message arrives on the same session (new turn).
      draft.updateContext(context);
    }
    return draft;
  }

  hasDraft(sessionId: string): boolean {
    return this.drafts.has(sessionId);
  }

  getDraft(sessionId: string): TeamsMessageDraft | undefined {
    return this.drafts.get(sessionId);
  }

  async finalize(sessionId: string, _context?: TurnContext, _isAssistant?: boolean): Promise<void> {
    const draft = this.drafts.get(sessionId);
    if (!draft) return;

    // Delete BEFORE awaiting to prevent concurrent finalize() calls
    // from double-finalizing the same draft (matches Telegram pattern)
    this.drafts.delete(sessionId);
    await draft.finalize();
  }

  cleanup(sessionId: string): void {
    this.drafts.delete(sessionId);
  }
}
