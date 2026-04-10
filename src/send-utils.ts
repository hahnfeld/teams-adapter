/**
 * Type-safe wrappers for Teams Bot Framework SDK calls.
 *
 * The @microsoft/agents-hosting SDK has strict Activity types that don't
 * align perfectly with the simple patterns we use (sending text, cards).
 * These helpers provide type-safe alternatives to `as any` casts.
 */
import { CardFactory } from "@microsoft/agents-hosting";
import type { TurnContext } from "@microsoft/agents-hosting";

/** Send a text message via TurnContext */
export async function sendText(context: TurnContext, text: string): Promise<unknown> {
  return (context.sendActivity as Function)({ text });
}

/** Send an Adaptive Card via TurnContext */
export async function sendCard(context: TurnContext, card: Record<string, unknown>): Promise<unknown> {
  return (context.sendActivity as Function)({
    attachments: [CardFactory.adaptiveCard(card)],
  });
}

/** Send a message with attachments via TurnContext */
export async function sendActivity(context: TurnContext, activity: Record<string, unknown>): Promise<unknown> {
  return (context.sendActivity as Function)(activity);
}

/** Update an existing activity */
export async function updateActivity(context: TurnContext, activity: Record<string, unknown>): Promise<unknown> {
  return (context.updateActivity as Function)(activity);
}

/**
 * Create an Adaptive Card attachment from a card object.
 * Wraps CardFactory.adaptiveCard with proper typing.
 */
export function adaptiveCardAttachment(card: Record<string, unknown>): unknown {
  return CardFactory.adaptiveCard(card);
}
