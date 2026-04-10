/**
 * Teams-specific validation functions for the install/configure wizard.
 * Mirrors the Telegram adapter's validators.ts pattern.
 */
import { log } from "@openacp/plugin-sdk";

/**
 * Validate bot credentials by acquiring a token from the MSA/AAD endpoint.
 * This proves the App ID and Password are correct and the bot registration exists.
 */
export async function validateBotCredentials(
  appId: string,
  appPassword: string,
  tenantId?: string,
): Promise<{ ok: true } | { ok: false; error: string }> {
  try {
    const tenant = tenantId || "botframework.com";
    const response = await fetch(
      `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "client_credentials",
          client_id: appId,
          client_secret: appPassword,
          scope: "https://api.botframework.com/.default",
        }).toString(),
      },
    );

    if (response.ok) {
      return { ok: true };
    }

    const data = (await response.json().catch(() => ({}))) as { error_description?: string; error?: string };
    if (data.error === "invalid_client") {
      return { ok: false, error: "Invalid App ID or Password. Check your Azure Bot registration." };
    }
    if (data.error === "unauthorized_client") {
      return { ok: false, error: "App ID is not authorized. Ensure it's registered as a Bot in Azure." };
    }
    return { ok: false, error: `Authentication failed (${response.status}). Verify your credentials.` };
  } catch (err) {
    return { ok: false, error: `Network error: ${(err as Error).message}` };
  }
}

/**
 * Validate a tenant ID by attempting token acquisition against it.
 * Also tries to resolve the tenant display name via the OpenID configuration.
 */
export async function validateTenant(
  appId: string,
  appPassword: string,
  tenantId: string,
): Promise<{ ok: true; tenantName?: string } | { ok: false; error: string }> {
  // First validate credentials work with this tenant
  const credResult = await validateBotCredentials(appId, appPassword, tenantId);
  if (!credResult.ok) {
    return { ok: false, error: `Tenant validation failed: ${credResult.error}` };
  }

  // Try to get tenant display name via OpenID config
  let tenantName: string | undefined;
  try {
    const oidcRes = await fetch(
      `https://login.microsoftonline.com/${tenantId}/v2.0/.well-known/openid-configuration`,
    );
    if (oidcRes.ok) {
      const oidc = (await oidcRes.json()) as { tenant_region_scope?: string; issuer?: string };
      // The issuer URL contains the tenant ID, confirming it's valid
      if (oidc.issuer?.includes(tenantId)) {
        tenantName = tenantId; // Display the GUID; tenant name requires Graph access
      }
    }
  } catch {
    // Non-critical — we already validated the credentials
  }

  return { ok: true, tenantName };
}

interface TeamInfo {
  id: string;
  name: string;
  channels: Array<{ id: string; name: string }>;
}

/**
 * Discover Teams and channels the bot has access to via Microsoft Graph API.
 * Requires the bot's app registration to have Team.ReadBasic.All or similar permissions.
 * Returns an empty array if Graph permissions aren't configured (graceful fallback).
 */
export async function discoverTeamsAndChannels(
  appId: string,
  appPassword: string,
  tenantId: string,
): Promise<{ ok: true; teams: TeamInfo[] } | { ok: false; error: string }> {
  try {
    // Get a Graph token
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "client_credentials",
          client_id: appId,
          client_secret: appPassword,
          scope: "https://graph.microsoft.com/.default",
        }).toString(),
      },
    );

    if (!tokenRes.ok) {
      return { ok: false, error: "Could not acquire Graph API token. Auto-discovery requires Graph permissions." };
    }

    const tokenData = (await tokenRes.json()) as { access_token: string };
    const token = tokenData.access_token;

    // List teams (requires Team.ReadBasic.All or TeamSettings.Read.All)
    const teamsRes = await fetch("https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName&$top=50", {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!teamsRes.ok) {
      if (teamsRes.status === 403) {
        return { ok: false, error: "Graph API permissions not configured. Add Team.ReadBasic.All to your app registration for auto-discovery." };
      }
      return { ok: false, error: `Graph API error (${teamsRes.status})` };
    }

    const teamsData = (await teamsRes.json()) as { value: Array<{ id: string; displayName: string }> };
    const teams: TeamInfo[] = [];

    for (const team of (teamsData.value ?? []).slice(0, 10)) {
      // Get channels for each team
      try {
        const channelsRes = await fetch(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/channels?$select=id,displayName&$top=25`,
          { headers: { Authorization: `Bearer ${token}` } },
        );
        const channelsData = channelsRes.ok
          ? ((await channelsRes.json()) as { value: Array<{ id: string; displayName: string }> })
          : { value: [] };

        teams.push({
          id: team.id,
          name: team.displayName,
          channels: channelsData.value.map((c) => ({ id: c.id, name: c.displayName })),
        });
      } catch {
        teams.push({ id: team.id, name: team.displayName, channels: [] });
      }
    }

    return { ok: true, teams };
  } catch (err) {
    return { ok: false, error: `Discovery failed: ${(err as Error).message}` };
  }
}

/**
 * Parse a Teams channel link to extract team and channel IDs.
 * Teams links look like:
 *   https://teams.microsoft.com/l/channel/19%3A...%40thread.tacv2/General?groupId=<teamId>&tenantId=<tenantId>
 */
export function parseTeamsLink(url: string): { teamId?: string; channelId?: string; tenantId?: string } {
  try {
    const parsed = new URL(url);
    const groupId = parsed.searchParams.get("groupId") ?? undefined;
    const tenantId = parsed.searchParams.get("tenantId") ?? undefined;

    // Channel ID is in the path: /l/channel/<encodedChannelId>/...
    let channelId: string | undefined;
    const pathMatch = parsed.pathname.match(/\/l\/channel\/([^/]+)/);
    if (pathMatch) {
      channelId = decodeURIComponent(pathMatch[1]);
    }

    return { teamId: groupId, channelId, tenantId };
  } catch {
    return {};
  }
}
