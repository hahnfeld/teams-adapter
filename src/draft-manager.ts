import type { TurnContext } from "@microsoft/agents-hosting";
import { log } from "@openacp/plugin-sdk";
import { splitMessage } from "./formatting.js";
import { sendText, updateActivity as updateTeamsActivity } from "./send-utils.js";

/** First flush fires quickly so the user sees content fast */
const FIRST_FLUSH_INTERVAL = 500;
/** Subsequent flushes are slower to avoid rate limits (Teams: 7 ops/sec/conversation) */
const UPDATE_FLUSH_INTERVAL = 3000;
const MAX_DISPLAY_LENGTH = 1900;

export interface MessageRef {
  activityId?: string;
  conversationId?: string;
}

export class TeamsMessageDraft {
  private buffer: string = "";
  private ref?: MessageRef;
  private flushTimer?: ReturnType<typeof setTimeout>;
  private flushPromise: Promise<void> = Promise.resolve();
  private lastSentBuffer: string = "";
  private displayTruncated = false;
  private firstFlushPending = false;
  private finalizing = false;

  constructor(
    private context: TurnContext,
    private sendQueue: { enqueue<T>(fn: () => Promise<T>, opts?: { type?: string }): Promise<T | undefined> },
    private sessionId: string,
  ) {}

  /** Update the TurnContext to the latest inbound turn (prevents stale context refs) */
  updateContext(context: TurnContext): void {
    this.context = context;
  }

  append(text: string): void {
    if (!text) return;
    this.buffer += text;
    this.scheduleFlush();
  }

  getBuffer(): string {
    return this.buffer;
  }

  private scheduleFlush(): void {
    if (this.flushTimer) return;
    // Fast first flush (500ms) for perceived responsiveness, then slower updates (3s)
    const interval = this.ref?.activityId ? UPDATE_FLUSH_INTERVAL : FIRST_FLUSH_INTERVAL;
    this.flushTimer = setTimeout(() => {
      this.flushTimer = undefined;
      this.flushPromise = this.flushPromise.then(() => this.flush()).catch(() => {});
    }, interval);
  }

  async flush(): Promise<void> {
    if (!this.buffer) return;
    if (this.firstFlushPending) return;

    const snapshot = this.buffer;

    let content = snapshot;
    let truncated = false;
    if (content.length > MAX_DISPLAY_LENGTH) {
      content = snapshot.slice(0, MAX_DISPLAY_LENGTH) + "…";
      truncated = true;
    }

    if (!content) return;

    if (!this.ref?.activityId) {
      this.firstFlushPending = true;
      try {
        const result = await this.sendQueue.enqueue(
          () => sendText(this.context, content) as Promise<unknown>,
          { type: "other" },
        );
        if (result) {
          const activityId = (result as { id?: string }).id;
          this.ref = {
            activityId,
            conversationId: this.context.activity.conversation?.id as string | undefined,
          };
          if (!truncated) {
            this.lastSentBuffer = snapshot;
            this.displayTruncated = false;
          } else {
            this.displayTruncated = true;
          }
        }
      } catch (err) {
        log.warn({ err, sessionId: this.sessionId }, "[TeamsMessageDraft] flush: sendActivity failed");
      } finally {
        this.firstFlushPending = false;
      }
    } else {
      if (!truncated && snapshot === this.lastSentBuffer) return;

      try {
        const result = await this.sendQueue.enqueue(
          () => updateTeamsActivity(this.context, {
            id: this.ref!.activityId,
            conversation: { id: this.ref!.conversationId },
            text: content,
          }) as Promise<unknown>,
          { type: "text" },
        );
        if (result !== undefined) {
          if (!truncated) {
            this.lastSentBuffer = snapshot;
            this.displayTruncated = false;
          } else {
            this.displayTruncated = true;
          }
        }
      } catch (err) {
        log.warn({ err, sessionId: this.sessionId, activityId: this.ref?.activityId }, "[TeamsMessageDraft] flush: updateActivity failed");
      }
    }
  }

  async stripPattern(pattern: RegExp): Promise<void> {
    if (!this.ref?.activityId || !this.buffer) return;

    let stripped: string;
    try {
      stripped = this.buffer.replace(pattern, "").trim();
    } catch (err) {
      log.warn({ err, sessionId: this.sessionId }, "[TeamsMessageDraft] stripPattern: replace failed");
      return;
    }

    if (stripped === this.buffer.trim()) return;

    if (!stripped) return;

    try {
      await this.sendQueue.enqueue(
        () => updateTeamsActivity(this.context, {
          id: this.ref!.activityId,
          conversation: { id: this.ref!.conversationId },
          text: stripped,
        }) as Promise<unknown>,
        { type: "other" },
      );
      // Only update state after successful send
      this.buffer = stripped;
      this.lastSentBuffer = stripped;
    } catch (err) {
      log.warn({ err, sessionId: this.sessionId, activityId: this.ref?.activityId }, "[TeamsMessageDraft] stripPattern: updateActivity failed");
    }
  }

  async finalize(): Promise<void> {
    if (this.finalizing) return;
    this.finalizing = true;
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

    await this.flushPromise;

    if (!this.buffer) return;

    if (this.ref?.activityId && this.buffer === this.lastSentBuffer && !this.displayTruncated) {
      return;
    }

    if (this.buffer.length <= MAX_DISPLAY_LENGTH) {
      const content = this.buffer;
      try {
        if (this.ref?.activityId) {
          await this.sendQueue.enqueue(
            () => updateTeamsActivity(this.context, {
              id: this.ref!.activityId,
              conversation: { id: this.ref!.conversationId },
              text: content,
            }) as Promise<unknown>,
            { type: "other" },
          );
        } else {
          await this.sendQueue.enqueue(
            () => sendText(this.context, content) as Promise<unknown>,
            { type: "other" },
          );
        }
        return;
      } catch {
        // Fall through to split approach
      }
    }

    const chunks = splitMessage(this.buffer, MAX_DISPLAY_LENGTH);

    for (let i = 0; i < chunks.length; i++) {
      const content = chunks[i];
      try {
        if (i === 0 && this.ref?.activityId) {
          await this.sendQueue.enqueue(
            () => updateTeamsActivity(this.context, {
              id: this.ref!.activityId,
              conversation: { id: this.ref!.conversationId },
              text: content,
            }) as Promise<unknown>,
            { type: "other" },
          );
        } else {
          const result = await this.sendQueue.enqueue(
            () => sendText(this.context, content) as Promise<unknown>,
            { type: "other" },
          );
          if (result && i === 0) {
            const activityId = (result as { id?: string }).id;
            this.ref = {
              activityId,
              conversationId: this.context.activity.conversation?.id as string | undefined,
            };
          }
        }
      } catch (err) {
        log.warn({ err, sessionId: this.sessionId, chunk: i }, "[TeamsMessageDraft] finalize: chunk send failed");
      }
    }
  }
}

export class TeamsDraftManager {
  private drafts = new Map<string, TeamsMessageDraft>();

  constructor(private sendQueue: { enqueue<T>(fn: () => Promise<T>, opts?: { type?: string }): Promise<T | undefined> }) {}

  getOrCreate(sessionId: string, context: TurnContext): TeamsMessageDraft {
    let draft = this.drafts.get(sessionId);
    if (!draft) {
      draft = new TeamsMessageDraft(context, this.sendQueue, sessionId);
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