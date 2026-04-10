/**
 * Persistent conversation reference store for proactive messaging.
 *
 * Teams bots can only send proactive messages if they have a stored
 * ConversationReference from a prior interaction. This store persists
 * references across bot restarts.
 *
 * References are captured on every inbound message and stored by:
 * - conversationId -> full reference (for replying to existing conversations)
 * - "service" -> the serviceUrl + credentials needed for createConversation
 *
 * @see https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/send-proactive-messages
 */
import { log } from "@openacp/plugin-sdk";
import * as fs from "node:fs";
import * as path from "node:path";

export interface StoredConversationReference {
  /** The Teams conversation ID */
  conversationId: string;
  /** Service URL for the bot connector */
  serviceUrl: string;
  /** Tenant ID */
  tenantId: string;
  /** Channel ID (Teams channel within a team) */
  channelId?: string;
  /** The bot's own ID in this conversation */
  botId: string;
  /** The bot's name */
  botName: string;
  /** Timestamp of last update */
  updatedAt: number;
}

export class ConversationStore {
  private references = new Map<string, StoredConversationReference>();
  private filePath: string;
  private dirty = false;
  private saveTimer?: ReturnType<typeof setInterval>;

  constructor(storageDir: string) {
    this.filePath = path.join(storageDir, "conversation-refs.json");
    this.load();

    // Auto-save every 30 seconds if dirty
    this.saveTimer = setInterval(() => {
      if (this.dirty) this.persist();
    }, 30_000);
    if (this.saveTimer.unref) this.saveTimer.unref();
  }

  /**
   * Store a conversation reference from an inbound activity.
   * Called on every inbound message to keep references fresh.
   */
  upsert(ref: StoredConversationReference): void {
    this.references.set(ref.conversationId, { ...ref, updatedAt: Date.now() });
    this.dirty = true;
  }

  /** Get a stored reference by conversation ID */
  get(conversationId: string): StoredConversationReference | undefined {
    return this.references.get(conversationId);
  }

  /** Get any stored reference (for proactive messaging when we don't know the conversation) */
  getAny(): StoredConversationReference | undefined {
    // Return the most recently updated reference
    let best: StoredConversationReference | undefined;
    for (const ref of this.references.values()) {
      if (!best || ref.updatedAt > best.updatedAt) best = ref;
    }
    return best;
  }

  /** Get all stored references */
  getAll(): StoredConversationReference[] {
    return Array.from(this.references.values());
  }

  destroy(): void {
    if (this.saveTimer) {
      clearInterval(this.saveTimer);
      this.saveTimer = undefined;
    }
    if (this.dirty) this.persist();
  }

  private load(): void {
    try {
      if (fs.existsSync(this.filePath)) {
        const data = JSON.parse(fs.readFileSync(this.filePath, "utf-8"));
        if (Array.isArray(data)) {
          for (const ref of data) {
            if (ref.conversationId) {
              this.references.set(ref.conversationId, ref);
            }
          }
        }
        log.info({ count: this.references.size }, "[ConversationStore] Loaded conversation references");
      }
    } catch (err) {
      log.warn({ err }, "[ConversationStore] Failed to load references");
    }
  }

  private persist(): void {
    try {
      const dir = path.dirname(this.filePath);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      // Atomic write: write to temp file then rename (prevents corruption on crash)
      const tmpPath = `${this.filePath}.tmp`;
      fs.writeFileSync(tmpPath, JSON.stringify(Array.from(this.references.values()), null, 2), { mode: 0o600 });
      fs.renameSync(tmpPath, this.filePath);
      this.dirty = false;
    } catch (err) {
      log.warn({ err }, "[ConversationStore] Failed to persist references");
    }
  }
}
