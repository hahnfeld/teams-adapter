import type { OutgoingMessage, NotificationMessage, DisplayVerbosity, ToolCallMeta, ToolUpdateMeta, PlanEntry, RenderedMessage } from "@openacp/plugin-sdk";
import { BaseRenderer } from "@openacp/plugin-sdk";
import { formatToolCall, formatToolUpdate, formatPlan, formatUsage } from "./formatting.js";
import type { AdaptiveCards } from "@microsoft/teams.cards";

/**
 * TeamsRenderer — renders messages using Adaptive Cards for rich formatting.
 * Extends BaseRenderer from plugin-sdk.
 */
export class TeamsRenderer extends BaseRenderer {
  renderToolCall(content: OutgoingMessage, verbosity: DisplayVerbosity): RenderedMessage {
    const meta = (content.metadata ?? {}) as Partial<ToolCallMeta>;
    return { body: formatToolCall(meta as ToolCallMeta, verbosity), format: "markdown" };
  }

  renderToolUpdate(content: OutgoingMessage, verbosity: DisplayVerbosity): RenderedMessage {
    const meta = (content.metadata ?? {}) as Partial<ToolUpdateMeta>;
    return { body: formatToolUpdate(meta as ToolUpdateMeta, verbosity), format: "markdown" };
  }

  renderPlan(content: OutgoingMessage): RenderedMessage {
    const entries = (content.metadata as { entries?: PlanEntry[] })?.entries ?? [];
    return { body: formatPlan(entries), format: "markdown" };
  }

  renderUsage(content: OutgoingMessage, verbosity: DisplayVerbosity): RenderedMessage {
    const meta = content.metadata as { tokensUsed?: number; contextSize?: number; cost?: number } | undefined;
    return { body: formatUsage(meta ?? {}, verbosity), format: "markdown" };
  }

  renderError(content: OutgoingMessage): RenderedMessage {
    return { body: `❌ **Error:** ${content.text}`, format: "markdown" };
  }

  renderNotification(notification: NotificationMessage): RenderedMessage {
    const emoji: Record<string, string> = {
      completed: "✅", error: "❌", permission: "🔐", input_required: "💬", budget_warning: "⚠️",
    };
    const icon = emoji[notification.type] || "ℹ️";
    const name = notification.sessionName ? ` **${notification.sessionName}**` : "";
    let text = `${icon}${name}: ${notification.summary}`;
    if (notification.deepLink) {
      text += `\n${notification.deepLink}`;
    }
    return { body: text, format: "markdown" };
  }

  renderSystemMessage(content: OutgoingMessage): RenderedMessage {
    return { body: content.text, format: "markdown" };
  }

  renderSessionEnd(_content: OutgoingMessage): RenderedMessage {
    return { body: "✅ Done", format: "markdown" };
  }

  renderModeChange(content: OutgoingMessage): RenderedMessage {
    const modeId = (content.metadata as Record<string, unknown>)?.modeId ?? "";
    return { body: `🔄 **Mode:** ${modeId}`, format: "markdown" };
  }

  renderConfigUpdate(): RenderedMessage {
    return { body: "⚙️ **Config updated**", format: "markdown" };
  }

  renderModelUpdate(content: OutgoingMessage): RenderedMessage {
    const modelId = (content.metadata as Record<string, unknown>)?.modelId ?? "";
    return { body: `🤖 **Model:** ${modelId}`, format: "markdown" };
  }
}

/**
 * Build an Adaptive Card from a markdown body string.
 * Wraps the text in a TextBlock with proper formatting.
 */
export function markdownToCard(text: string, color?: string): AdaptiveCards.AdaptiveCard {
  return {
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text,
        wrap: true,
        ...(color ? { color: color as AdaptiveCards.TextBlockColor } : {}),
      },
    ],
  };
}

/**
 * Build an Adaptive Card notification.
 */
export function buildNotificationCard(notification: NotificationMessage): AdaptiveCards.AdaptiveCard {
  const emoji: Record<string, string> = {
    completed: "✅", error: "❌", permission: "🔐", input_required: "💬", budget_warning: "⚠️",
  };
  const icon = emoji[notification.type] || "ℹ️";
  const name = notification.sessionName ? ` **${notification.sessionName}**` : "";
  const text = `${icon}${name}: ${notification.summary}`;

  return {
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text,
        wrap: true,
        size: "Medium",
        weight: "Bolder",
      },
    ],
  };
}