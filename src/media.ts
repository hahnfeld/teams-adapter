import type { Attachment } from "@openacp/plugin-sdk";

export async function downloadTeamsFile(
  url: string,
  fileName: string,
): Promise<Buffer | null> {
  // TODO: Implement Teams file download via Graph API or Teams API
  // Similar to Discord's downloadDiscordAttachment in media.ts
  throw new Error("Not yet implemented");
}

export function isAttachmentTooLarge(size: number): boolean {
  // Teams file size limit is 250MB (compared to Discord's 25MB)
  return size > 250 * 1024 * 1024;
}