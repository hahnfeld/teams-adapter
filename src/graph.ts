/**
 * Microsoft Graph API client for Teams file operations.
 *
 * Provides authenticated file download (from Teams attachment URLs) and
 * file upload (to OneDrive/SharePoint for sharing via adaptive cards).
 *
 * Uses the Microsoft Graph REST API v1.0:
 * - Download: GET /drives/{driveId}/items/{itemId}/content
 * - Upload (small, <4MB): PUT /me/drive/root:/{path}:/content
 * - Upload (large, >4MB): POST /me/drive/root:/{path}:/createUploadSession
 *
 * Authentication uses client credentials flow (app-only) with the bot's
 * Azure AD app registration. Requires Files.ReadWrite.All or Sites.ReadWrite.All
 * application permissions for upload, and User.Read for basic operations.
 *
 * @see https://learn.microsoft.com/en-us/graph/api/driveitem-put-content
 * @see https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession
 */
import { log } from "@openacp/plugin-sdk";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const SMALL_FILE_LIMIT = 4 * 1024 * 1024; // 4MB — Graph simple upload limit

interface GraphTokenCache {
  token: string;
  expiresAt: number;
}

/**
 * Manages Graph API authentication and file operations for a Teams bot.
 *
 * Requires Azure AD app registration with:
 * - Application permission: Files.ReadWrite.All (for upload to SharePoint/OneDrive)
 * - Delegated or Application permission for downloading Teams attachments
 */
export class GraphFileClient {
  private tokenCache: GraphTokenCache | null = null;

  constructor(
    private tenantId: string,
    private clientId: string,
    private clientSecret: string,
  ) {}

  /**
   * Acquire an access token via client credentials flow (app-only).
   * Caches the token and refreshes 5 minutes before expiry.
   */
  private async getToken(): Promise<string> {
    if (this.tokenCache && Date.now() < this.tokenCache.expiresAt - 300_000) {
      return this.tokenCache.token;
    }

    const tokenUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      grant_type: "client_credentials",
      client_id: this.clientId,
      client_secret: this.clientSecret,
      scope: "https://graph.microsoft.com/.default",
    });

    const response = await fetch(tokenUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    if (!response.ok) {
      // Sanitize error — Azure AD error responses may contain client_id and diagnostic info
      throw new Error(`Graph token acquisition failed (HTTP ${response.status})`);
    }

    const data = (await response.json()) as { access_token: string; expires_in: number };
    this.tokenCache = {
      token: data.access_token,
      expiresAt: Date.now() + data.expires_in * 1000,
    };

    return this.tokenCache.token;
  }

  /**
   * Download a file from a Teams attachment content URL with proper authentication.
   *
   * Teams attachment URLs (e.g., https://teams.microsoft.com/api/...) require
   * a Bearer token. This method fetches the file with the bot's Graph token.
   */
  async downloadFile(contentUrl: string, maxSize: number): Promise<Buffer | null> {
    if (!contentUrl) return null;

    try {
      const token = await this.getToken();
      const response = await fetch(contentUrl, {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (!response.ok) {
        log.warn(
          { url: contentUrl.slice(0, 80), status: response.status },
          "[GraphFileClient] Download failed",
        );
        return null;
      }

      const contentLength = response.headers.get("content-length");
      if (contentLength && parseInt(contentLength, 10) > maxSize) {
        log.warn({ contentLength, maxSize }, "[GraphFileClient] File exceeds size limit");
        return null;
      }

      const arrayBuffer = await response.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);

      if (buffer.length > maxSize) {
        log.warn({ size: buffer.length, maxSize }, "[GraphFileClient] File exceeds size limit after download");
        return null;
      }

      return buffer;
    } catch (err) {
      log.warn({ err }, "[GraphFileClient] Download error");
      return null;
    }
  }

  /**
   * Upload a file to the bot's OneDrive for Business and return a sharing URL.
   *
   * For files <= 4MB, uses simple PUT upload.
   * For files > 4MB, uses upload session (resumable upload).
   *
   * The file is uploaded to /openacp-files/{sessionId}/{fileName} to keep
   * files organized by session.
   *
   * @returns A sharing link URL that Teams users can access, or null on failure.
   *
   * @see https://learn.microsoft.com/en-us/graph/api/driveitem-put-content
   * @see https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession
   */
  async uploadFile(
    sessionId: string,
    fileName: string,
    content: Buffer,
    contentType: string,
  ): Promise<string | null> {
    try {
      const token = await this.getToken();
      const remotePath = `/openacp-files/${sessionId}/${fileName}`;

      let driveItemId: string;

      if (content.length <= SMALL_FILE_LIMIT) {
        // Simple upload for small files (≤4MB)
        driveItemId = await this.simpleUpload(token, remotePath, content, contentType);
      } else {
        // Resumable upload for large files (>4MB, up to 250MB)
        driveItemId = await this.resumableUpload(token, remotePath, content);
      }

      // Create a sharing link
      const shareUrl = await this.createSharingLink(token, driveItemId);
      return shareUrl;
    } catch (err) {
      log.error({ err, sessionId, fileName }, "[GraphFileClient] Upload failed");
      return null;
    }
  }

  /**
   * Simple PUT upload for files ≤ 4MB.
   * PUT /me/drive/root:/{path}:/content
   */
  private async simpleUpload(
    token: string,
    remotePath: string,
    content: Buffer,
    contentType: string,
  ): Promise<string> {
    const url = `${GRAPH_BASE}/me/drive/root:${remotePath}:/content`;
    const response = await fetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": contentType,
      },
      body: content,
    });

    if (!response.ok) {
      throw new Error(`Simple upload failed (HTTP ${response.status})`);
    }

    const item = (await response.json()) as { id: string };
    return item.id;
  }

  /**
   * Resumable upload for files > 4MB.
   * Creates an upload session and uploads in 10MB chunks.
   *
   * @see https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession
   */
  private async resumableUpload(
    token: string,
    remotePath: string,
    content: Buffer,
  ): Promise<string> {
    // Step 1: Create upload session
    const sessionUrl = `${GRAPH_BASE}/me/drive/root:${remotePath}:/createUploadSession`;
    const sessionResponse = await fetch(sessionUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        item: { "@microsoft.graph.conflictBehavior": "rename" },
      }),
    });

    if (!sessionResponse.ok) {
      throw new Error(`Create upload session failed (HTTP ${sessionResponse.status})`);
    }

    const session = (await sessionResponse.json()) as { uploadUrl: string };
    const uploadUrl = session.uploadUrl;

    // Step 2: Upload in 10MB chunks
    const chunkSize = 10 * 1024 * 1024; // 10MB — must be multiple of 320KB
    const totalSize = content.length;
    let offset = 0;
    let lastResponse: Response | null = null;

    while (offset < totalSize) {
      const end = Math.min(offset + chunkSize, totalSize);
      const chunk = content.subarray(offset, end);

      // Upload session URL is pre-authenticated — Authorization header must NOT be
      // included per Graph docs (it causes a 401 if present on the session URL).
      lastResponse = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": String(chunk.length),
          "Content-Range": `bytes ${offset}-${end - 1}/${totalSize}`,
        },
        body: chunk,
      });

      if (!lastResponse.ok && lastResponse.status !== 202) {
        throw new Error(`Chunk upload failed at offset ${offset} (HTTP ${lastResponse.status})`);
      }

      offset = end;
    }

    if (!lastResponse) throw new Error("No upload response received");

    // Final chunk should return 200 or 201 with the drive item body.
    // A 202 on the final chunk means the server hasn't finalized yet (rare edge case).
    if (lastResponse.status !== 200 && lastResponse.status !== 201) {
      throw new Error(`Upload completed but final response was HTTP ${lastResponse.status} (expected 200/201)`);
    }

    const item = (await lastResponse.json()) as { id?: string };
    if (!item.id) {
      throw new Error("Upload response missing drive item ID");
    }
    return item.id;
  }

  /**
   * Create an anonymous sharing link for a drive item.
   * POST /me/drive/items/{id}/createLink
   */
  private async createSharingLink(token: string, driveItemId: string): Promise<string> {
    const url = `${GRAPH_BASE}/me/drive/items/${driveItemId}/createLink`;
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: "view",
        scope: "organization",
      }),
    });

    if (!response.ok) {
      // Fallback: return a direct webUrl instead
      const itemUrl = `${GRAPH_BASE}/me/drive/items/${driveItemId}`;
      const itemResponse = await fetch(itemUrl, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (itemResponse.ok) {
        const item = (await itemResponse.json()) as { webUrl?: string };
        if (item.webUrl) return item.webUrl;
      }
      throw new Error(`createLink failed (${response.status})`);
    }

    const link = (await response.json()) as { link: { webUrl: string } };
    return link.link.webUrl;
  }
}
