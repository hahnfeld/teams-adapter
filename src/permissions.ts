import type { TurnContext } from "@microsoft/agents-hosting";
import { nanoid } from "nanoid";
import type { PermissionRequest, NotificationMessage, Session } from "@openacp/plugin-sdk";
import { log } from "@openacp/plugin-sdk";
import { sendCard, sendText, updateActivity, adaptiveCardAttachment } from "./send-utils.js";

interface PendingPermission {
  sessionId: string;
  requestId: string;
  options: { id: string; isAllow: boolean }[];
  activityId?: string;
  conversationId?: string;
  createdAt: number;
}

/** Try to extract a tool/action name from the permission description text */
function parseDescriptionContext(description: string): { tool?: string; target?: string; risk: "high" | "normal" } {
  let tool: string | undefined;
  let target: string | undefined;
  let risk: "high" | "normal" = "normal";

  // Common patterns: "Allow Read /path/to/file", "Execute: bash command", "Edit src/foo.ts"
  const toolMatch = description.match(/^(Read|Write|Edit|Execute|Bash|Delete|Install|Command|Fetch|Search|Agent)\b/i);
  if (toolMatch) tool = toolMatch[1];

  // Extract file paths
  const pathMatch = description.match(/[`"]?([/~][\w./\-@]+|[a-zA-Z]:\\[\w\\.\-]+)[`"]?/);
  if (pathMatch) target = pathMatch[1];

  // High-risk keywords
  if (/\b(delete|remove|drop|rm\b|force|sudo|--hard|--force|push)\b/i.test(description)) {
    risk = "high";
  }

  return { tool, target, risk };
}

export class PermissionHandler {
  private pending: Map<string, PendingPermission> = new Map();
  private pendingTimestamps: Map<string, number> = new Map();

  constructor(
    private getSession: (sessionId: string) => Session | undefined,
    private sendNotification: (notification: NotificationMessage) => Promise<void>,
  ) {}

  private evictStale(): void {
    const MAX_PENDING = 100;
    if (this.pending.size < MAX_PENDING) return;
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
      options: request.options.map((o) => ({ id: o.id, isAllow: o.isAllow })),
      activityId: context.activity.id as string | undefined,
      conversationId: context.activity.conversation?.id as string | undefined,
      createdAt: now,
    });
    this.pendingTimestamps.set(callbackKey, now);

    const { tool, target, risk } = parseDescriptionContext(request.description);
    const headerColor = risk === "high" ? "Attention" : "Warning";

    // Build rich permission card (Adaptive Card v1.2 for mobile compatibility)
    const body: unknown[] = [
      // Header with risk-colored icon
      {
        type: "ColumnSet",
        columns: [
          { type: "Column", width: "auto", items: [{ type: "TextBlock", text: risk === "high" ? "⚠️" : "🔐", size: "Large" }] },
          {
            type: "Column", width: "stretch", items: [
              { type: "TextBlock", text: "Permission Request", weight: "Bolder", size: "Medium", color: headerColor },
              ...(session.name ? [{ type: "TextBlock", text: session.name, size: "Small", isSubtle: true, spacing: "None" }] : []),
            ],
          },
        ],
      },
      // Description
      { type: "TextBlock", text: request.description, wrap: true, spacing: "Medium" },
    ];

    // Context facts (tool, target)
    const facts: Array<{ title: string; value: string }> = [];
    if (tool) facts.push({ title: "Action", value: tool });
    if (target) facts.push({ title: "Target", value: target });
    if (facts.length > 0) {
      body.push({ type: "FactSet", facts, spacing: "Small" });
    }

    // Risk warning for destructive operations
    if (risk === "high") {
      body.push({
        type: "TextBlock",
        text: "⚠️ This action may be destructive or hard to reverse.",
        color: "Attention",
        size: "Small",
        wrap: true,
        spacing: "Medium",
      });
    }

    const card = {
      type: "AdaptiveCard" as const,
      version: "1.2" as const,
      body,
      // Action.Submit sends activity.value as the flat data object
      actions: request.options.map((option) => ({
        type: "Action.Submit" as const,
        title: `${option.isAllow ? "✅" : "❌"} ${option.label}`,
        ...(risk === "high" && !option.isAllow ? { style: "destructive" } : {}),
        data: { verb: option.isAllow ? "allow" : "deny", sessionId: session.id, callbackKey, requestId: request.id },
      })),
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

    void this.sendNotification({
      sessionId: session.id,
      sessionName: session.name,
      type: "permission",
      summary: request.description,
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
      // Update the original card to show expired state
      try {
        await sendText(context, "❌ Permission request expired or already responded to.");
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

    // Update the original card to show the response with who/when
    if (pending.activityId && pending.conversationId) {
      try {
        const decision = verb === "always" ? "Always Allowed" : verb === "allow" ? "Allowed" : "Denied";
        const decisionColor = verb === "deny" ? "Attention" : "Good";
        const decisionIcon = verb === "deny" ? "❌" : "✅";

        const updatedCard = {
          type: "AdaptiveCard" as const,
          version: "1.2" as const,
          body: [
            {
              type: "ColumnSet",
              columns: [
                { type: "Column", width: "auto", items: [{ type: "TextBlock", text: decisionIcon, size: "Large" }] },
                {
                  type: "Column", width: "stretch", items: [
                    { type: "TextBlock", text: `Permission — ${decision}`, weight: "Bolder", color: decisionColor },
                  ],
                },
              ],
            },
            {
              type: "FactSet",
              facts: [
                { title: "Responded by", value: respondedBy },
                { title: "Response time", value: elapsed < 60 ? `${elapsed}s` : `${Math.round(elapsed / 60)}m` },
              ],
              spacing: "Small",
            },
          ],
          actions: [] as unknown[],
        };
        await updateActivity(context, {
          id: pending.activityId,
          conversation: { id: pending.conversationId },
          attachments: [adaptiveCardAttachment(updatedCard as Record<string, unknown>)],
        });
      } catch { /* ignore update failures */ }
    }

    return true;
  }
}
