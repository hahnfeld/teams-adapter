import type { TurnContext } from "@microsoft/agents-hosting";
import { log } from "@openacp/plugin-sdk";
import { splitMessage } from "./formatting.js";
import { sendText } from "./send-utils.js";

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
    // Streaming updates via updateActivity don't work with the teams.apps SDK
    // context (it only has send(), not updateActivity()). Instead of trying to
    // update in-place, we skip periodic flushes and let finalize() send the
    // complete message. The typing indicator keeps the user informed.
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

    if (!this.buffer) return;

    // Send the complete text as new message(s), split if needed.
    // We don't use updateActivity because the teams.apps SDK context
    // doesn't support it — only send() is available.
    const chunks = splitMessage(this.buffer, MAX_DISPLAY_LENGTH);

    for (let i = 0; i < chunks.length; i++) {
      const content = chunks[i];
      try {
        await this.sendQueue.enqueue(
          () => sendText(this.context, content) as Promise<unknown>,
          { type: "other" },
        );
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