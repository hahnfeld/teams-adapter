import type { TurnContext } from "@microsoft/agents-hosting";
import { nanoid } from "nanoid";
import type { PermissionRequest, NotificationMessage, Session } from "@openacp/plugin-sdk";
import { log } from "@openacp/plugin-sdk";
import type { SessionMessageManager } from "./message-composer.js";

interface PendingPermission {
  sessionId: string;
  requestId: string;
  description: string;
  entryId: string;
  options: { id: string; label: string; isAllow: boolean }[];
  createdAt: number;
}

// ─── PermissionHandler ────────────────────────────────────────────────────────

export class PermissionHandler {
  private pending: Map<string, PendingPermission> = new Map();
  private pendingTimestamps: Map<string, number> = new Map();
  private _evictionTimer: ReturnType<typeof setInterval> | undefined;

  constructor(
    private getSession: (sessionId: string) => Session | undefined,
    private sendNotification: (notification: NotificationMessage) => Promise<void>,
    private composer: SessionMessageManager,
  ) {
    this._evictionTimer = setInterval(() => this.evictStale(true), 5 * 60 * 1000);
    this._evictionTimer.unref();
  }

  dispose(): void {
    if (this._evictionTimer) {
      clearInterval(this._evictionTimer);
      this._evictionTimer = undefined;
    }
    this.pending.clear();
    this.pendingTimestamps.clear();
  }

  private evictStale(force = false): void {
    const MAX_PENDING = 100;
    if (!force && this.pending.size < MAX_PENDING) return;
    if (this.pending.size === 0) return;
    const now = Date.now();
    const staleThreshold = 10 * 60 * 1000;
    for (const [key, ts] of this.pendingTimestamps) {
      if (now - ts > staleThreshold) {
        this.pending.delete(key);
        this.pendingTimestamps.delete(key);
      }
    }
  }

  async sendPermissionRequest(
    session: Session,
    request: PermissionRequest,
    context: TurnContext,
  ): Promise<void> {
    this.evictStale();
    const callbackKey = nanoid(8);
    const now = Date.now();

    // Build Action.Submit buttons for the card
    const actions = request.options.map((option) => ({
      type: "Action.Submit" as const,
      title: `${option.isAllow ? "✅" : "❌"} ${option.label}`,
      data: { verb: option.isAllow ? "allow" : "deny", sessionId: session.id, callbackKey, requestId: request.id },
    }));

    // Add to the existing session card (even if sealed after usage), or create fresh
    const msg = this.composer.get(session.id) ?? this.composer.getOrCreate(session.id, context);
    const entryId = msg.addPermission(request.description, actions);

    this.pending.set(callbackKey, {
      sessionId: session.id,
      requestId: request.id,
      description: request.description,
      entryId,
      options: request.options.map((o) => ({ id: o.id, label: o.label, isAllow: o.isAllow })),
      createdAt: now,
    });
    this.pendingTimestamps.set(callbackKey, now);

    this.sendNotification({
      sessionId: session.id,
      sessionName: session.name,
      type: "permission",
      summary: request.description,
    }).catch((err) => {
      log.warn({ err }, "[PermissionHandler] Notification failed");
    });
  }

  async handleCardAction(
    context: TurnContext,
    verb: string,
    sessionId: string,
    callbackKey: string,
    requestId: string,
  ): Promise<boolean> {
    if (verb !== "allow" && verb !== "deny" && verb !== "always") return false;

    const pending = this.pending.get(callbackKey);
    if (!pending) {
      log.debug({ callbackKey }, "[PermissionHandler] Permission expired or already responded to");
      return true;
    }

    const session = this.getSession(pending.sessionId);
    const respondedBy = context.activity.from?.name ?? context.activity.from?.id ?? "Unknown";
    const elapsed = Math.round((Date.now() - pending.createdAt) / 1000);

    log.info({ requestId: pending.requestId, verb, sessionId, respondedBy, elapsedSec: elapsed }, "[PermissionHandler] Permission responded");

    if (session?.permissionGate?.requestId === pending.requestId) {
      const wantAllow = verb === "allow" || verb === "always";
      const option = pending.options.find((o) => o.isAllow === wantAllow);
      if (option) {
        session.permissionGate.resolve(option.id);
      } else {
        const fallback = pending.options[wantAllow ? 0 : pending.options.length - 1];
        log.warn({ requestId: pending.requestId, verb, optionCount: pending.options.length }, "[PermissionHandler] No matching option, using fallback");
        session.permissionGate.resolve(fallback?.id ?? pending.options[0]?.id ?? "denied");
      }
    }

    // Update the permission entry in the session card
    const decision = verb === "always" ? "Always Allowed" : verb === "allow" ? "Allowed" : "Denied";
    const elapsedStr = elapsed < 60 ? `${elapsed}s` : `${Math.round(elapsed / 60)}m`;
    const msg = this.composer.get(pending.sessionId);
    if (msg) {
      msg.resolvePermission(pending.entryId, `${decision} (${respondedBy}, ${elapsedStr})`);
    }

    this.pending.delete(callbackKey);
    this.pendingTimestamps.delete(callbackKey);

    return true;
  }
}
