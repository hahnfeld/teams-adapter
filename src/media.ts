import type { Attachment } from "@openacp/plugin-sdk";
import { log } from "@openacp/plugin-sdk";

const MAX_FILE_SIZE = 250 * 1024 * 1024; // 250MB Teams limit

export async function downloadTeamsFile(
  url: string,
  fileName: string,
): Promise<Buffer | null> {
  if (!url) return null;

  try {
    const response = await fetch(url);
    if (!response.ok) {
      log.warn({ url, fileName, status: response.status, statusText: response.statusText }, "[media] downloadTeamsFile: HTTP error");
      return null;
    }

    // Check Content-Length header before buffering the full file
    const contentLength = response.headers.get("content-length");
    if (contentLength) {
      const size = parseInt(contentLength, 10);
      if (size > MAX_FILE_SIZE) {
        throw new Error(`File too large: ${size} bytes (max ${MAX_FILE_SIZE})`);
      }
    }

    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    if (buffer.length > MAX_FILE_SIZE) {
      throw new Error(`File too large: ${buffer.length} bytes (max ${MAX_FILE_SIZE})`);
    }

    return buffer;
  } catch (err) {
    // Return null for transient errors, throw only for truly unexpected issues
    if (err instanceof Error && err.message.startsWith("File too large")) {
      throw err;
    }
    // Attachment fetch failures are non-fatal — return null and let the message proceed
    return null;
  }
}

export function isAttachmentTooLarge(size: number): boolean {
  // Teams file size limit is 250MB (compared to Discord's 25MB)
  return size > MAX_FILE_SIZE;
}

export function buildFileAttachmentCard(
  fileName: string,
  size: number,
  contentUrl: string,
  contentType: string,
): { type: "AdaptiveCard"; version: "1.4"; body: unknown[]; actions: unknown[] } {
  return {
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      { type: "TextBlock", text: `📎 ${fileName}`, weight: "Bolder", wrap: true },
      {
        type: "TextBlock",
        text: `Size: ${(size / 1024 / 1024).toFixed(1)} MB`,
        size: "Small",
        isSubtle: true,
      },
    ],
    actions: [
      {
        type: "Action.OpenUrl",
        title: "Download",
        url: contentUrl,
        role: "button",
      },
    ],
  };
}
