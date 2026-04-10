import { TurnContext } from "@microsoft/teams.botbuilder";
import { splitMessage } from "./formatting.js";
import type { SendQueue } from "@openacp/plugin-sdk";

const FLUSH_INTERVAL = 5000;
const MAX_DISPLAY_LENGTH = 1900;

export interface MessageRef {
  activityId?: string;
  conversationId?: string;
}

/**
 * Teams-specific message draft that batches text updates and sends
 * them as a single Teams message, editing it in place via updateActivity().
 */
export class TeamsMessageDraft {
  private buffer: string = "";
  private ref?: MessageRef;
  private flushTimer?: ReturnType<typeof setTimeout>;
  private flushPromise: Promise<void> = Promise.resolve();
  private lastSentBuffer: string = "";
  private displayTruncated = false;
  private firstFlushPending = false;

  constructor(
    private context: TurnContext,
    private sendQueue: SendQueue,
    private sessionId: string,
  ) {}

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
    this.flushTimer = setTimeout(() => {
      this.flushTimer = undefined;
      this.flushPromise = this.flushPromise
        .then(() => this.flush())
        .catch(() => {});
    }, FLUSH_INTERVAL);
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
          () => this.context.sendActivity({ text: content }),
          { type: "other" },
        );
        if (result) {
          const activityId = (result as { id?: string }).id;
          this.ref = {
            activityId,
            conversationId: this.context.activity.conversation?.id,
          };
          if (!truncated) {
            this.lastSentBuffer = snapshot;
            this.displayTruncated = false;
          } else {
            this.displayTruncated = true;
          }
        }
      } catch {
        // send failed — next flush will retry
      } finally {
        this.firstFlushPending = false;
      }
    } else {
      if (!truncated && snapshot === this.lastSentBuffer) return;

      try {
        const result = await this.sendQueue.enqueue(
          () => this.context.updateActivity({
            id: this.ref.activityId,
            conversation: { id: this.ref.conversationId },
            text: content,
          }),
          { type: "text", key: this.sessionId },
        );
        if (result !== undefined) {
          if (!truncated) {
            this.lastSentBuffer = snapshot;
            this.displayTruncated = false;
          } else {
            this.displayTruncated = true;
          }
        }
      } catch {
        // Don't reset ref — transient errors should not cause duplicate sends
      }
    }
  }

  async stripPattern(pattern: RegExp): Promise<void> {
    if (!this.ref?.activityId || !this.buffer) return;

    const stripped = this.buffer.replace(pattern, "").trim();
    if (stripped === this.buffer.trim()) return;

    this.buffer = stripped;
    this.lastSentBuffer = stripped;

    if (!stripped) return;

    try {
      await this.sendQueue.enqueue(
        () => this.context.updateActivity({
          id: this.ref.activityId,
          conversation: { id: this.ref.conversationId },
          text: stripped,
        }),
        { type: "other" },
      );
    } catch {
      // Best effort — non-critical edit
    }
  }

  async finalize(): Promise<void> {
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
            () => this.context.updateActivity({
              id: this.ref.activityId,
              conversation: { id: this.ref.conversationId },
              text: content,
            }),
            { type: "other" },
          );
        } else {
          await this.sendQueue.enqueue(
            () => this.context.sendActivity({ text: content }),
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
            () => this.context.updateActivity({
              id: this.ref.activityId,
              conversation: { id: this.ref.conversationId },
              text: content,
            }),
            { type: "other" },
          );
        } else {
          const result = await this.sendQueue.enqueue(
            () => this.context.sendActivity({ text: content }),
            { type: "other" },
          );
          if (result && i === 0) {
            const activityId = (result as { id?: string }).id;
            this.ref = {
              activityId,
              conversationId: this.context.activity.conversation?.id,
            };
          }
        }
      } catch {
        // Skip this chunk — best effort
      }
    }
  }
}

/**
 * Teams-specific draft manager.
 * Batches text updates into drafts before sending, handles TTS strip pattern.
 * Draft message stored per session, flushed on tool call or timeout.
 */
export class TeamsDraftManager {
  private drafts = new Map<string, TeamsMessageDraft>();
  private textBuffers = new Map<string, string>();

  constructor(private sendQueue: SendQueue) {}

  getOrCreate(sessionId: string, context: TurnContext): TeamsMessageDraft {
    let draft = this.drafts.get(sessionId);
    if (!draft) {
      draft = new TeamsMessageDraft(context, this.sendQueue, sessionId);
      this.drafts.set(sessionId, draft);
    }
    return draft;
  }

  hasDraft(sessionId: string): boolean {
    return this.drafts.has(sessionId);
  }

  getDraft(sessionId: string): TeamsMessageDraft | undefined {
    return this.drafts.get(sessionId);
  }

  appendText(sessionId: string, text: string): void {
    this.textBuffers.set(sessionId, (this.textBuffers.get(sessionId) ?? "") + text);
  }

  async finalize(sessionId: string, context?: TurnContext, isAssistant?: boolean): Promise<void> {
    const draft = this.drafts.get(sessionId);
    if (!draft) return;

    this.drafts.delete(sessionId);
    await draft.finalize();

    if (isAssistant && context) {
      const fullText = this.textBuffers.get(sessionId);
      this.textBuffers.delete(sessionId);
      if (fullText) {
        // TODO: Detect action patterns and send action buttons as follow-up
        // Similar to Discord's detectAction/storeAction/buildActionKeyboard
      }
    } else {
      this.textBuffers.delete(sessionId);
    }
  }

  cleanup(sessionId: string): void {
    this.drafts.delete(sessionId);
    this.textBuffers.delete(sessionId);
  }
}