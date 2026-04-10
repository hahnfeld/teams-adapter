export interface TeamsChannelConfig {
  enabled: boolean;
  botAppId: string;
  botAppPassword: string;
  tenantId: string;
  teamId: string;
  channelId: string;
  notificationChannelId: string | null;
  assistantThreadId: string | null;
  /** Azure AD client secret for Graph API file operations. Optional — falls back to card-only display. */
  graphClientSecret?: string;
}

export interface TeamsPlatformData {
  teamId: string;
  channelId: string;
  threadId?: string;
  messageId?: string;
  skillMsgId?: string;
}