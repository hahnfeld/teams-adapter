import type { TurnContext } from "@microsoft/agents-hosting";
import { nanoid } from "nanoid";
import type { PermissionRequest, NotificationMessage, Session } from "@openacp/plugin-sdk";
import { log } from "@openacp/plugin-sdk";
import { sendCard, updateActivity, adaptiveCardAttachment } from "./send-utils.js";
import { buildLevel1, buildLevel2, escapeMd } from "./message-composer.js";

interface PendingPermission {
  sessionId: string;
  requestId: string;
  description: string;
  options: { id: string; label: string; isAllow: boolean }[];
  activityId?: string;
  conversationId?: string;
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
    this.pending.set(callbackKey, {
      sessionId: session.id,
      requestId: request.id,
      description: request.description,
      options: request.options.map((o) => ({ id: o.id, label: o.label, isAllow: o.isAllow })),
      activityId: context.activity.id as string | undefined,
      conversationId: context.activity.conversation?.id as string | undefined,
      createdAt: now,
    });
    this.pendingTimestamps.set(callbackKey, now);

    // Build permission card matching the info Container style
    const card = {
      type: "AdaptiveCard" as const,
      version: "1.4" as const,
      body: [
        {
          type: "Container",
          spacing: "Small",
          items: [
            buildLevel1("🔐", "Permission"),
            buildLevel2(request.description),
            // Option buttons — inline ActionSet (smaller than top-level actions)
            {
              type: "ActionSet",
              spacing: "Small",
              actions: request.options.map((option) => ({
                type: "Action.Submit" as const,
                title: `${option.isAllow ? "✅" : "❌"} ${option.label}`,
                data: { verb: option.isAllow ? "allow" : "deny", sessionId: session.id, callbackKey, requestId: request.id },
              })),
            },
          ],
        },
      ],
      width: "stretch",
    };

    try {
      const result = await sendCard(context, card as Record<string, unknown>);
      const activityId = (result as { id?: string })?.id;
      const pendingEntry = this.pending.get(callbackKey);
      if (pendingEntry && activityId) {
        pendingEntry.activityId = activityId;
        pendingEntry.conversationId = context.activity.conversation?.id as string | undefined;
      }
    } catch (err) {
      log.warn({ err, sessionId: session.id }, "[PermissionHandler] Failed to send permission request");
      return;
    }

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
      try {
        // Show expired as an info-style card
        const card = {
          type: "AdaptiveCard",
          version: "1.4",
          body: [{
            type: "Container",
            spacing: "Small",
            items: [
              buildLevel1("🔐", "Permission"),
              buildLevel2("Expired or already responded to"),
            ],
          }],
          width: "stretch",
        };
        await sendCard(context, card);
      } catch (err) {
        log.warn({ err, callbackKey }, "[PermissionHandler] Failed to send expired message");
      }
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

    this.pending.delete(callbackKey);
    this.pendingTimestamps.delete(callbackKey);

    // Update the original card to show the resolved state
    if (pending.activityId && pending.conversationId) {
      try {
        const decision = verb === "always" ? "Always Allowed" : verb === "allow" ? "Allowed" : "Denied";
        const decisionIcon = verb === "deny" ? "❌" : "✅";
        const elapsedStr = elapsed < 60 ? `${elapsed}s` : `${Math.round(elapsed / 60)}m`;

        const updatedCard = {
          type: "AdaptiveCard" as const,
          version: "1.4" as const,
          body: [{
            type: "Container",
            spacing: "Small",
            items: [
              buildLevel1(decisionIcon, `Permission — ${escapeMd(decision)}`),
              buildLevel2(`${pending.description} (${respondedBy}, ${elapsedStr})`),
            ],
          }],
          width: "stretch",
        };
        await updateActivity(context, {
          id: pending.activityId,
          attachments: [adaptiveCardAttachment(updatedCard as Record<string, unknown>)],
        });
      } catch { /* ignore update failures */ }
    }

    return true;
  }
}
