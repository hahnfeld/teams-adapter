/**
 * Type definitions bridging the Microsoft Teams SDK with our adapter.
 *
 * The @microsoft/teams.apps and @microsoft/agents-hosting SDKs have
 * incomplete TypeScript types in several areas. These interfaces provide
 * type-safe wrappers for the patterns we use, eliminating `as any` casts.
 */

/** Activity sent to or from the bot */
export interface TeamsActivity {
  type: string;
  id?: string;
  text?: string;
  from?: { id: string; name?: string; aadObjectId?: string };
  conversation?: { id: string; tenantId?: string; conversationType?: string };
  channelId?: string;
  serviceUrl?: string;
  attachments?: TeamsAttachment[];
  value?: Record<string, unknown>;
  replyToId?: string;
}

/** Attachment in a Teams message */
export interface TeamsAttachment {
  contentType: string;
  contentUrl?: string;
  name?: string;
  content?: unknown;
}

/** Card action from an Adaptive Card button */
export interface TeamsCardAction {
  verb: string;
  data: Record<string, unknown>;
}

/** Result from sendActivity */
export interface ActivityResult {
  id?: string;
}

/** Adaptive Card structure (v1.2 for mobile compat) */
export interface AdaptiveCard {
  type: "AdaptiveCard";
  version: "1.2";
  body: AdaptiveCardElement[];
  actions?: AdaptiveCardAction[];
}

export type AdaptiveCardElement =
  | TextBlock
  | ColumnSet
  | Column
  | Container;

export interface TextBlock {
  type: "TextBlock";
  text: string;
  weight?: "Default" | "Bolder" | "Lighter";
  size?: "Default" | "Small" | "Medium" | "Large" | "ExtraLarge";
  color?: "Default" | "Dark" | "Light" | "Accent" | "Good" | "Warning" | "Attention";
  wrap?: boolean;
  isSubtle?: boolean;
  spacing?: "None" | "Small" | "Default" | "Medium" | "Large" | "ExtraLarge" | "Padding";
}

export interface ColumnSet {
  type: "ColumnSet";
  columns: Column[];
}

export interface Column {
  type: "Column";
  width: "auto" | "stretch" | string;
  items: AdaptiveCardElement[];
}

export interface Container {
  type: "Container";
  items: AdaptiveCardElement[];
}

export type AdaptiveCardAction =
  | ActionExecute
  | ActionOpenUrl
  | ActionSubmit;

export interface ActionExecute {
  type: "Action.Execute";
  title: string;
  data: Record<string, unknown>;
}

export interface ActionOpenUrl {
  type: "Action.OpenUrl";
  title: string;
  url: string;
}

export interface ActionSubmit {
  type: "Action.Submit";
  title: string;
  data: Record<string, unknown>;
}

/**
 * Helper to create a type-safe Adaptive Card.
 * Avoids the `as any` cast when passing to CardFactory.adaptiveCard().
 */
export function createAdaptiveCard(
  body: AdaptiveCardElement[],
  actions?: AdaptiveCardAction[],
): AdaptiveCard {
  return {
    type: "AdaptiveCard",
    version: "1.2",
    body,
    actions,
  };
}

/** Create a text block */
export function textBlock(text: string, opts?: Partial<TextBlock>): TextBlock {
  return { type: "TextBlock", text, wrap: true, ...opts };
}

/** Create an Action.Submit button (v1.2 compatible — use instead of Action.Execute for mobile) */
export function actionSubmit(title: string, data: Record<string, unknown>): ActionSubmit {
  return { type: "Action.Submit", title, data };
}

/** Create an Action.OpenUrl button */
export function actionOpenUrl(title: string, url: string): ActionOpenUrl {
  return { type: "Action.OpenUrl", title, url };
}
