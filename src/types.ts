/** Default Bot Framework HTTP server port. */
export const DEFAULT_BOT_PORT = 3978;

export interface TeamsChannelConfig {
  enabled: boolean;
  botAppId: string;
  botAppPassword: string;
  tenantId: string;
  teamId: string;
  channelId: string;
  notificationChannelId: string | null;
  assistantThreadId: string | null;
  /** Port for the Bot Framework HTTP server (default: 3978). Separate from the OpenACP API port. */
  botPort?: number;
  /** Azure AD client secret for Graph API file operations. Optional — falls back to card-only display. */
  graphClientSecret?: string;
  /** Tunnel method: "devtunnel" (recommended), "builtin" (auto-create via tunnel service), or "manual". Default: "devtunnel". */
  tunnelMethod?: "devtunnel" | "builtin" | "manual";
}

export interface TeamsPlatformData {
  teamId: string;
  channelId: string;
  threadId?: string;
  messageId?: string;
  skillMsgId?: string;
}