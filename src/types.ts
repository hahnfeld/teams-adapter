import type { IChannelAdapter } from "@openacp/plugin-sdk";

export interface TeamsChannelConfig extends Partial<IChannelAdapter> {
  enabled: boolean;
  botAppId: string;
  botAppPassword: string;
  tenantId: string;
  teamId: string;
  channelId: string;
  notificationChannelId: string | null;
  assistantThreadId: string | null;
}

export interface TeamsPlatformData {
  teamId: string;
  channelId: string;
  threadId?: string;
  messageId?: string;
  skillMsgId?: string;
}

export interface CommandsAssistantContext {
  threadId: string;
  getSession: () => import("@openacp/plugin-sdk").Session | null;
  respawn: () => Promise<void>;
}