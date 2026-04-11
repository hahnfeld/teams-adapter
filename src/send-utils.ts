/**
 * Type-safe wrappers for sending messages via Teams SDK contexts.
 *
 * The @microsoft/teams.apps SDK uses a different context API than the
 * Bot Framework TurnContext. These helpers detect which context type
 * is available and call the appropriate method:
 *   - teams.apps context: context.send({ text }) / context.reply({ text })
 *   - Bot Framework TurnContext: context.sendActivity({ text })
 */
import { CardFactory } from "@microsoft/agents-hosting";
import type { TurnContext } from "@microsoft/agents-hosting";

/**
 * Normalize newlines for Teams rendering.
 * Teams collapses single \n in markdown — use \n\n for line breaks.
 * Avoids doubling already-doubled newlines.
 */
function teamsNewlines(text: string): string {
  return text.replace(/(?<!\n)\n(?!\n)/g, "\n\n");
}

/** Send a text message via context (supports both TurnContext and teams.apps context) */
export async function sendText(context: TurnContext, text: string): Promise<unknown> {
  const activity = { type: "message", text: teamsNewlines(text), textFormat: "markdown" };
  if (typeof (context as any).send === "function") {
    return (context as any).send(activity);
  }
  return (context.sendActivity as Function)(activity);
}

/** Send an Adaptive Card via context */
export async function sendCard(context: TurnContext, card: Record<string, unknown>): Promise<unknown> {
  const activity = {
    type: "message",
    attachments: [CardFactory.adaptiveCard(card)],
  };
  if (typeof (context as any).send === "function") {
    return (context as any).send(activity);
  }
  return (context.sendActivity as Function)(activity);
}

/** Send a message with attachments via context */
export async function sendActivity(context: TurnContext, activity: Record<string, unknown>): Promise<unknown> {
  const merged: Record<string, unknown> = { type: "message", textFormat: "markdown", ...activity };
  if (typeof merged.text === "string") {
    merged.text = teamsNewlines(merged.text);
  }
  if (typeof (context as any).send === "function") {
    return (context as any).send(merged);
  }
  return (context.sendActivity as Function)(merged);
}

/** Update an existing activity */
export async function updateActivity(context: TurnContext, activity: Record<string, unknown>): Promise<unknown> {
  if (typeof (context as any).updateActivity === "function") {
    return (context.updateActivity as Function)(activity);
  }
  // teams.apps SDK doesn't have updateActivity on context — skip silently
  return undefined;
}

/**
 * Create an Adaptive Card attachment from a card object.
 * Wraps CardFactory.adaptiveCard with proper typing.
 */
export function adaptiveCardAttachment(card: Record<string, unknown>): unknown {
  return CardFactory.adaptiveCard(card);
}
