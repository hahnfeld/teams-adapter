import { CardFactory } from "@microsoft/teams.cards";
import type { TurnContext } from "@microsoft/teams.botbuilder";
import { nanoid } from "nanoid";
import type { PermissionRequest, NotificationMessage, Session } from "@openacp/plugin-sdk";
import { log } from "@openacp/plugin-sdk";

interface PendingPermission {
  sessionId: string;
  requestId: string;
  options: { id: string; isAllow: boolean }[];
  activityId?: string;
  conversationId?: string;
}

export class PermissionHandler {
  private pending: Map<string, PendingPermission> = new Map();

  constructor(
    private getSession: (sessionId: string) => Session | undefined,
    private sendNotification: (notification: NotificationMessage) => Promise<void>,
  ) {}

  async sendPermissionRequest(
    session: Session,
    request: PermissionRequest,
    context: TurnContext,
  ): Promise<void> {
    const callbackKey = nanoid(8);
    this.pending.set(callbackKey, {
      sessionId: session.id,
      requestId: request.id,
      options: request.options.map((o) => ({ id: o.id, isAllow: o.isAllow })),
      activityId: context.activity.id,
      conversationId: context.activity.conversation?.id,
    });

    const card = {
      type: "AdaptiveCard" as const,
      version: "1.4" as const,
      body: [
        {
          type: "TextBlock",
          text: "🔐 Permission Request",
          weight: "Bolder" as const,
          color: "Warning" as const,
          size: "Medium" as const,
        },
        {
          type: "TextBlock",
          text: request.description,
          wrap: true,
          spacing: "Medium" as const,
        },
      ],
      actions: request.options.map((option) => ({
        type: "Action.Execute" as const,
        title: `${option.isAllow ? "✅" : "❌"} ${option.label}`,
        data: { verb: option.isAllow ? "allow" : "deny", sessionId: session.id, callbackKey, requestId: request.id },
      })),
    };

    try {
      const result = await context.sendActivity({
        attachments: [CardFactory.adaptiveCard(card)],
      });
      const activityId = (result as { id?: string }).id;
      const pendingEntry = this.pending.get(callbackKey);
      if (pendingEntry && activityId) {
        pendingEntry.activityId = activityId;
        pendingEntry.conversationId = context.activity.conversation?.id;
      }
    } catch (err) {
      log.warn({ err, sessionId: session.id }, "[PermissionHandler] Failed to send permission request");
      return;
    }

    void this.sendNotification({
      sessionId: session.id,
      sessionName: session.name,
      type: "permission",
      summary: request.description,
    });
  }

  async handleCardAction(context: TurnContext, verb: string, sessionId: string, callbackKey: string, requestId: string): Promise<boolean> {
    if (verb !== "allow" && verb !== "deny" && verb !== "always") return false;

    const pending = this.pending.get(callbackKey);
    if (!pending) {
      try {
        await context.sendActivity({ text: "❌ Permission request expired" });
      } catch { /* ignore */ }
      return true;
    }

    const session = this.getSession(pending.sessionId);

    log.info(
      { requestId: pending.requestId, verb, sessionId },
      "[PermissionHandler] Permission responded",
    );

    if (session?.permissionGate?.requestId === pending.requestId) {
      const option = verb === "allow" || verb === "always"
        ? pending.options.find((o) => o.isAllow)
        : pending.options.find((o) => !o.isAllow);
      session.permissionGate.resolve(option?.id ?? "");
    }

    this.pending.delete(callbackKey);

    try {
      await context.sendActivity({ text: `✅ Permission ${verb === "always" ? "always-allowed" : verb}d` });
    } catch { /* ignore */ }

    // Remove action buttons from the original card by editing it
    if (pending.activityId && pending.conversationId) {
      try {
        const updatedCard = {
          type: "AdaptiveCard" as const,
          version: "1.4" as const,
          body: [
            {
              type: "TextBlock",
              text: "🔐 Permission Request — Responded",
              weight: "Bolder" as const,
              isSubtle: true,
            },
            {
              type: "TextBlock",
              text: `✅ ${verb === "always" ? "Always Allowed" : verb === "allow" ? "Allowed" : "Denied"}`,
              wrap: true,
            },
          ],
          actions: [],
        };
        await context.updateActivity({
          id: pending.activityId,
          conversation: { id: pending.conversationId },
          attachments: [CardFactory.adaptiveCard(updatedCard)],
        });
      } catch { /* ignore */ }
    }

    return true;
  }
}