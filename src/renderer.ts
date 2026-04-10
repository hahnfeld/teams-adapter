import type { OutgoingMessage, NotificationMessage, DisplayVerbosity, ToolCallMeta, ToolUpdateMeta, PlanEntry } from "@openacp/plugin-sdk";
import { BaseRenderer } from "@openacp/plugin-sdk";
import { formatToolCall, formatToolUpdate, formatPlan, formatUsage } from "./formatting.js";

export class TeamsRenderer extends BaseRenderer {
  renderToolCall(content: OutgoingMessage, verbosity: DisplayVerbosity): { body: string; format: "markdown" } {
    const meta = (content.metadata ?? {}) as Partial<ToolCallMeta>;
    return { body: formatToolCall(meta as ToolCallMeta, verbosity), format: "markdown" };
  }

  renderToolUpdate(content: OutgoingMessage, verbosity: DisplayVerbosity): { body: string; format: "markdown" } {
    const meta = (content.metadata ?? {}) as Partial<ToolUpdateMeta>;
    return { body: formatToolUpdate(meta as ToolUpdateMeta, verbosity), format: "markdown" };
  }

  renderPlan(content: OutgoingMessage): { body: string; format: "markdown" } {
    const entries = (content.metadata as { entries?: PlanEntry[] })?.entries ?? [];
    return { body: formatPlan(entries), format: "markdown" };
  }

  renderUsage(content: OutgoingMessage, verbosity: DisplayVerbosity): { body: string; format: "markdown" } {
    const meta = content.metadata as { tokensUsed?: number; contextSize?: number; cost?: number } | undefined;
    return { body: formatUsage(meta ?? {}, verbosity), format: "markdown" };
  }

  renderError(content: OutgoingMessage): { body: string; format: "markdown" } {
    return { body: `❌ **Error:** ${content.text}`, format: "markdown" };
  }

  renderNotification(notification: NotificationMessage): { body: string; format: "markdown" } {
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

  renderSystemMessage(content: OutgoingMessage): { body: string; format: "markdown" } {
    return { body: content.text, format: "markdown" };
  }

  renderSessionEnd(_content: OutgoingMessage): { body: string; format: "markdown" } {
    return { body: "✅ Done", format: "markdown" };
  }

  renderModeChange(content: OutgoingMessage): { body: string; format: "markdown" } {
    const modeId = (content.metadata as Record<string, unknown>)?.modeId ?? "";
    return { body: `🔄 **Mode:** ${modeId}`, format: "markdown" };
  }

  renderConfigUpdate(): { body: string; format: "markdown" } {
    return { body: "⚙️ **Config updated**", format: "markdown" };
  }

  renderModelUpdate(content: OutgoingMessage): { body: string; format: "markdown" } {
    const modelId = (content.metadata as Record<string, unknown>)?.modelId ?? "";
    return { body: `🤖 **Model:** ${modelId}`, format: "markdown" };
  }
}