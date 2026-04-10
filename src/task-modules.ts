/**
 * Task Module (dialog) definitions for Teams.
 *
 * Task Modules render as modal popups containing Adaptive Cards.
 * They're triggered by Action.Submit with msteams.type = "task/fetch",
 * and the bot handles task/fetch and task/submit invoke activities.
 *
 * @see https://learn.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/task-modules/task-modules-bots
 */

/** Build the "New Session" dialog card — agent selection + workspace input */
export function buildNewSessionDialog(
  agents: Array<{ name: string; description?: string }>,
  defaultWorkspace: string,
): { task: unknown } {
  const agentChoices = agents.map((a) => ({
    title: a.name,
    value: a.name,
  }));

  // Default to first agent if available
  if (agentChoices.length === 0) {
    agentChoices.push({ title: "openacp", value: "openacp" });
  }

  const card = {
    type: "AdaptiveCard",
    version: "1.2",
    body: [
      { type: "TextBlock", text: "New Session", weight: "Bolder", size: "Large" },
      { type: "TextBlock", text: "Select an agent and workspace to start a new coding session.", wrap: true, isSubtle: true, spacing: "Small" },

      // Agent selection
      { type: "TextBlock", text: "Agent", weight: "Bolder", spacing: "Large" },
      {
        type: "Input.ChoiceSet",
        id: "agent",
        style: "compact",
        value: agentChoices[0].value,
        choices: agentChoices,
      },

      // Workspace input
      { type: "TextBlock", text: "Workspace (project directory)", weight: "Bolder", spacing: "Large" },
      {
        type: "Input.Text",
        id: "workspace",
        placeholder: "/path/to/project",
        value: defaultWorkspace,
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Create Session",
        data: { dialogAction: "new-session" },
      },
    ],
  };

  return {
    task: {
      type: "continue",
      value: {
        title: "New Session",
        height: "medium",
        width: "medium",
        card: { contentType: "application/vnd.microsoft.card.adaptive", content: card },
      },
    },
  };
}

/** Build the "Settings" dialog card — view and edit key settings */
export function buildSettingsDialog(
  settings: {
    defaultAgent?: string;
    workspace?: string;
    outputMode?: string;
    sessionId?: string;
    sessionAgent?: string;
    sessionModel?: string;
    sessionBypass?: boolean;
    sessionTts?: string;
  },
): { task: unknown } {
  const card = {
    type: "AdaptiveCard",
    version: "1.2",
    body: [
      { type: "TextBlock", text: "Settings", weight: "Bolder", size: "Large" },

      // Global settings
      { type: "TextBlock", text: "Global", weight: "Bolder", spacing: "Large", separator: true },
      {
        type: "Input.ChoiceSet",
        id: "outputMode",
        label: "Output Detail Level",
        style: "compact",
        value: settings.outputMode ?? "medium",
        choices: [
          { title: "Low — minimal output", value: "low" },
          { title: "Medium — balanced (default)", value: "medium" },
          { title: "High — verbose with details", value: "high" },
        ],
      },

      // Session settings (if in a session)
      ...(settings.sessionAgent ? [
        { type: "TextBlock", text: "Current Session", weight: "Bolder", spacing: "Large", separator: true },
        {
          type: "FactSet",
          facts: [
            { title: "Agent", value: settings.sessionAgent },
            ...(settings.sessionModel ? [{ title: "Model", value: settings.sessionModel }] : []),
            { title: "Bypass", value: settings.sessionBypass ? "On" : "Off" },
            ...(settings.sessionTts ? [{ title: "TTS", value: settings.sessionTts }] : []),
          ],
        },
        {
          type: "Input.Toggle",
          id: "bypass",
          title: "Auto-approve permissions",
          value: settings.sessionBypass ? "true" : "false",
          valueOn: "true",
          valueOff: "false",
        },
      ] : []),
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Save",
        data: { dialogAction: "save-settings", ...(settings.sessionId ? { sessionId: settings.sessionId } : {}) },
      },
    ],
  };

  return {
    task: {
      type: "continue",
      value: {
        title: "Settings",
        height: "medium",
        width: "medium",
        card: { contentType: "application/vnd.microsoft.card.adaptive", content: card },
      },
    },
  };
}

/** Build a simple message response to close the dialog */
export function buildDialogMessage(text: string): { task: unknown } {
  return {
    task: {
      type: "message",
      value: text,
    },
  };
}
