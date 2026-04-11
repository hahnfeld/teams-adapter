import type { OpenACPPlugin, InstallContext, OpenACPCore } from "@openacp/plugin-sdk";
import { log } from "@openacp/plugin-sdk";
import { TeamsAdapter } from "./adapter.js";
import type { TeamsChannelConfig } from "./types.js";
import { DEFAULT_BOT_PORT } from "./types.js";

/**
 * Factory for the Teams adapter plugin.
 *
 * Includes a full interactive `install()` wizard that guides users through:
 * 1. Azure Bot registration (with Portal URLs and step-by-step instructions)
 * 2. Credential validation (real-time token acquisition test)
 * 3. Tenant configuration (single vs multi-tenant)
 * 4. Team/channel selection (auto-discovery via Graph or manual)
 * 5. Optional notification channel and Graph API file sharing
 */
export default function createTeamsPlugin(): OpenACPPlugin {
  let adapter: TeamsAdapter | null = null;

  return {
    name: "@hahnfeld/teams-adapter",
    version: "1.0.0",
    description: "Microsoft Teams adapter with Adaptive Cards, commands, and streaming",
    essential: false,
    permissions: ["services:register", "kernel:access", "events:read", "commands:register"],
    // TODO: Add Zod settingsSchema when @openacp/plugin-sdk exports a schema builder.
    // Required fields: enabled, botAppId, botAppPassword, tenantId, teamId, channelId
    // Optional: notificationChannelId, assistantThreadId, graphClientSecret

    // ─── Interactive Install Wizard ──────────────────────────────────────

    async install(ctx: InstallContext) {
      const { terminal, settings } = ctx;

      const { validateBotCredentials, validateTenant, discoverTeamsAndChannels, parseTeamsLink } =
        await import("./validators.js");

      // ── Step 1: Azure Bot Registration Guidance ──

      terminal.note(
        "This wizard will help you connect OpenACP to Microsoft Teams.\n" +
        "You'll need an Azure Bot registration. If you don't have one yet,\n" +
        "follow these steps first:\n" +
        "\n" +
        "  1. Go to: https://portal.azure.com/#create/Microsoft.AzureBot\n" +
        "  2. Fill in:\n" +
        "     - Bot handle: any unique name (e.g. 'openacp-bot')\n" +
        "     - Pricing: Free (F0) for testing\n" +
        "     - App type: 'Single Tenant' for enterprise, 'Multi Tenant' for public\n" +
        "     - Creation type: 'Create new Microsoft App ID'\n" +
        "  3. Click 'Create' and wait for deployment\n" +
        "  4. Go to the Bot resource → Settings → Configuration\n" +
        "     - Copy the 'Microsoft App ID'\n" +
        "  5. Go to 'Manage Password' → 'New client secret'\n" +
        "     - Copy the secret value immediately (it's shown only once)\n" +
        "  6. Under 'Channels', add the 'Microsoft Teams' channel\n" +
        "\n" +
        "Docs: https://learn.microsoft.com/en-us/azure/bot-service/bot-service-quickstart-registration",
        "Azure Bot Setup",
      );

      const ready = await terminal.confirm({
        message: "Do you have your Bot App ID and Password ready?",
        initialValue: true,
      });
      if (!ready) {
        terminal.log.info("No worries! Set up your Azure Bot first, then run this again.");
        terminal.cancel("Setup cancelled — re-run when ready.");
        return;
      }

      // ── Step 2: Bot App ID ──

      let botAppId = await terminal.text({
        message: "Bot App ID (Microsoft App ID from Azure Portal):",
        placeholder: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        validate: (val) => {
          const trimmed = val.trim();
          if (!trimmed) return "App ID cannot be empty";
          if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(trimmed)) {
            return "App ID should be a GUID (e.g. 12345678-1234-1234-1234-123456789abc)";
          }
          return undefined;
        },
      });
      botAppId = botAppId.trim();

      // ── Step 3: Bot App Password ──

      terminal.log.info("Find this in Azure Portal → Bot resource → Manage Password → Client secrets");

      let botAppPassword = await terminal.password({
        message: "Bot App Password (client secret):",
        validate: (val) => {
          if (!val.trim()) return "Password cannot be empty";
          return undefined;
        },
      });
      botAppPassword = botAppPassword.trim();

      // ── Step 4: Tenant Configuration ──
      // Collected before credential validation so we can use the correct tenant endpoint.

      terminal.log.info("");
      terminal.note(
        "Azure bots can be single-tenant (one organization) or multi-tenant (any org).\n" +
        "\n" +
        "  Single-tenant: For enterprise use within your organization.\n" +
        "                 Find your Tenant ID at:\n" +
        "                 Azure Portal → Microsoft Entra ID → Overview → Tenant ID\n" +
        "\n" +
        "  Multi-tenant:  For bots available to any Microsoft 365 organization.\n" +
        "                 Uses the default 'botframework.com' tenant.",
        "Tenant Type",
      );

      const tenantType = await terminal.select({
        message: "What type of bot registration?",
        options: [
          { value: "single", label: "Single-tenant (enterprise)", hint: "Most common for internal bots" },
          { value: "multi", label: "Multi-tenant (public)" },
        ],
      });

      let tenantId = "botframework.com";
      if (tenantType === "single") {
        while (true) {
          const tenantInput = await terminal.text({
            message: "Tenant ID (GUID from Azure Portal → Entra ID → Overview):",
            placeholder: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            validate: (val) => {
              const trimmed = val.trim();
              if (!trimmed) return "Tenant ID cannot be empty";
              if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(trimmed)) {
                return "Tenant ID should be a GUID";
              }
              return undefined;
            },
          });
          tenantId = tenantInput.trim();

          const spin = terminal.spinner();
          spin.start("Validating tenant...");
          const result = await validateTenant(botAppId, botAppPassword, tenantId);
          if (result.ok) {
            spin.stop(`Tenant validated: ${result.tenantName ?? tenantId}`);
            break;
          }
          spin.fail(result.error);

          const action = await terminal.select({
            message: "What to do?",
            options: [
              { value: "retry", label: "Re-enter tenant ID" },
              { value: "skip", label: "Use as-is (skip validation)" },
            ],
          });
          if (action === "skip") break;
        }
      }

      // ── Step 4b: Validate Credentials (now that we know the tenant) ──

      let credentialsValidated = false;
      while (!credentialsValidated) {
        const spin = terminal.spinner();
        spin.start("Validating bot credentials...");
        const result = await validateBotCredentials(botAppId, botAppPassword, tenantId !== "botframework.com" ? tenantId : undefined);
        if (result.ok) {
          spin.stop("Bot credentials validated successfully");
          credentialsValidated = true;
          break;
        }
        spin.fail(result.error);

        const action = await terminal.select({
          message: "What would you like to do?",
          options: [
            { value: "retry", label: "Re-enter App ID and password" },
            { value: "skip", label: "Skip validation (use as-is)" },
          ],
        });
        if (action === "skip") break;

        botAppId = await terminal.text({
          message: "Bot App ID:",
          defaultValue: botAppId,
          validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
        });
        botAppId = botAppId.trim();

        botAppPassword = await terminal.password({
          message: "Bot App Password:",
          validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
        });
        botAppPassword = botAppPassword.trim();
      }

      // ── Step 5: Team & Channel Selection ──

      terminal.log.info("");

      let teamId = "";
      let channelId = "";

      // Try auto-discovery (only for single-tenant bots — multi-tenant can't use Graph with botframework.com)
      let discovery: Awaited<ReturnType<typeof discoverTeamsAndChannels>> = { ok: false, error: "skipped" };
      if (tenantType === "single") {
        const spin2 = terminal.spinner();
        spin2.start("Discovering Teams and channels...");
        discovery = await discoverTeamsAndChannels(botAppId, botAppPassword, tenantId);

        if (discovery.ok && discovery.teams.length > 0) {
          spin2.stop(`Found ${discovery.teams.length} team(s)`);

          const teamOptions = discovery.teams.map((t) => ({
            value: t.id,
            label: t.name,
            hint: `${t.channels.length} channel(s)`,
          }));
          teamOptions.push({ value: "__manual__", label: "Enter manually instead", hint: "" });

          const selectedTeam = await terminal.select({
            message: "Which team should the bot operate in?",
            options: teamOptions,
          });

          if (selectedTeam !== "__manual__") {
            teamId = selectedTeam;
            const team = discovery.teams.find((t) => t.id === selectedTeam);

            if (team && team.channels.length > 0) {
              const channelOptions = team.channels.map((c) => ({
                value: c.id,
                label: c.name,
              }));
              channelId = await terminal.select({
                message: "Which channel should be the default?",
                options: channelOptions,
              });
            } else {
              channelId = await terminal.text({
                message: "Channel ID (no channels found — enter manually):",
                validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
              });
              channelId = channelId.trim();
            }
          }
        } else {
          if (discovery.ok) {
            spin2.stop("No teams found (bot may not be added to any team yet)");
          } else {
            spin2.stop("Auto-discovery not available — will use manual entry");
            terminal.log.info(`(${discovery.error})`);
          }
        }
      } // end single-tenant discovery

      // Manual entry fallback
      if (!teamId || !channelId) {
        terminal.log.info("");
        terminal.note(
          "To find your Team and Channel IDs:\n" +
          "\n" +
          "  Option A — From a channel link:\n" +
          "    1. Open Microsoft Teams\n" +
          "    2. Right-click the channel name → 'Get link to channel'\n" +
          "    3. Paste the link below — we'll extract the IDs automatically\n" +
          "\n" +
          "  Option B — Manual entry:\n" +
          "    Team ID = the 'groupId' parameter from the link\n" +
          "    Channel ID = the encoded string in the path (e.g. 19:xxx@thread.tacv2)",
          "Finding Team & Channel IDs",
        );

        const method = await terminal.select({
          message: "How to provide Team and Channel IDs?",
          options: [
            { value: "link", label: "Paste a channel link (easiest)", hint: "Right-click channel → Get link" },
            { value: "manual", label: "Enter IDs manually" },
          ],
        });

        if (method === "link") {
          const link = await terminal.text({
            message: "Paste the Teams channel link:",
            validate: (v) => {
              if (!v.trim()) return "Link cannot be empty";
              if (!v.includes("teams.microsoft.com") && !v.includes("teams.cloud.microsoft")) return "This doesn't look like a Teams link";
              return undefined;
            },
          });

          const parsed = parseTeamsLink(link.trim());
          if (parsed.teamId && parsed.channelId) {
            teamId = parsed.teamId;
            channelId = parsed.channelId;
            terminal.log.success(`Extracted Team ID: ${teamId}`);
            terminal.log.success(`Extracted Channel ID: ${channelId}`);
            if (parsed.tenantId && tenantType === "single" && parsed.tenantId !== tenantId) {
              terminal.log.warning(`Link tenant (${parsed.tenantId}) differs from configured tenant (${tenantId})`);
            }
          } else {
            terminal.log.warning("Could not extract IDs from link. Please enter manually.");
          }
        }

        if (!teamId) {
          teamId = await terminal.text({
            message: "Team ID (groupId GUID):",
            placeholder: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
          });
          teamId = teamId.trim();
        }

        if (!channelId) {
          channelId = await terminal.text({
            message: "Channel ID (e.g. 19:abc123@thread.tacv2):",
            validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
          });
          channelId = channelId.trim();
        }
      }

      // ── Step 6: Notification Channel (Optional) ──

      terminal.log.info("");
      const wantNotifications = await terminal.confirm({
        message: "Set up a dedicated notification channel? (Session completions, errors, permission alerts)",
        initialValue: false,
      });

      let notificationChannelId: string | null = null;
      if (wantNotifications) {
        if (discovery.ok) {
          const team = discovery.teams.find((t) => t.id === teamId);
          if (team && team.channels.length > 1) {
            const otherChannels = team.channels
              .filter((c) => c.id !== channelId)
              .map((c) => ({ value: c.id, label: c.name }));
            otherChannels.push({ value: "__manual__", label: "Enter manually" });

            const selected = await terminal.select({
              message: "Which channel for notifications?",
              options: otherChannels,
            });
            if (selected !== "__manual__") {
              notificationChannelId = selected;
            }
          }
        }

        if (!notificationChannelId) {
          const nid = await terminal.text({
            message: "Notification channel ID (or leave empty to skip):",
            defaultValue: "",
          });
          notificationChannelId = nid.trim() || null;
        }
      }

      // ── Step 7: Graph API for File Sharing (Optional) ──

      terminal.log.info("");
      const wantGraph = await terminal.confirm({
        message: "Enable file sharing via OneDrive? (Allows sharing agent-generated files in Teams)",
        initialValue: false,
      });

      let graphClientSecret: string | undefined;
      if (wantGraph) {
        terminal.note(
          "File sharing requires a Graph API client secret with Files.ReadWrite.All permission.\n" +
          "\n" +
          "To set this up:\n" +
          "  1. Azure Portal → App Registrations → find your bot's app\n" +
          "  2. API Permissions → Add → Microsoft Graph → Application permissions\n" +
          "     → Files.ReadWrite.All → Grant admin consent\n" +
          "  3. Certificates & secrets → New client secret → copy the value\n" +
          "\n" +
          "Note: You can use the same app registration as your bot.\n" +
          "The client secret can be different from the bot password.",
          "Graph API Setup",
        );

        const useExisting = await terminal.confirm({
          message: "Use the same client secret as the bot password?",
          initialValue: true,
        });

        if (useExisting) {
          graphClientSecret = botAppPassword;
          terminal.log.success("Using bot password for Graph API");
        } else {
          graphClientSecret = await terminal.password({
            message: "Graph API client secret:",
            validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
          });
          graphClientSecret = graphClientSecret.trim();
        }
      }

      // ── Step 8: Save & Summary ──

      // ── Step 8b: Bot Port ──

      terminal.log.info("");
      terminal.note(
        "The Teams adapter runs its own HTTP server for Bot Framework messages.\n" +
        "This is separate from the OpenACP API server.\n" +
        "\n" +
        `  Default port: ${DEFAULT_BOT_PORT} (Bot Framework standard)\n` +
        "  Your tunnel must point to this port, not the OpenACP API port.",
        "Bot Port",
      );

      const useDefaultPort = await terminal.confirm({
        message: `Use the default bot port (${DEFAULT_BOT_PORT})?`,
        initialValue: true,
      });

      let botPort = DEFAULT_BOT_PORT;
      if (!useDefaultPort) {
        const portInput = await terminal.text({
          message: "Bot port:",
          defaultValue: String(DEFAULT_BOT_PORT),
          validate: (v) => {
            const n = Number(v.trim());
            if (isNaN(n) || !Number.isInteger(n) || n < 1 || n > 65535) return "Port must be 1–65535";
            return undefined;
          },
        });
        botPort = Number(portInput.trim());
      }

      // ── Step 8c: Tunnel Method ──

      terminal.log.info("");
      terminal.note(
        "Azure Bot Service needs a public URL to reach your bot.\n" +
        "A tunnel exposes your local bot port to the internet.\n" +
        "\n" +
        `  Your bot port: ${botPort}`,
        "Tunnel Setup",
      );

      const tunnelMethod = await terminal.select({
        message: "How should the bot port be tunneled?",
        options: [
          { value: "devtunnel", label: "@hahnfeld/devtunnel-provider (Recommended)", hint: "Microsoft Dev Tunnels — persistent URLs" },
          { value: "builtin", label: "Built-in tunnel service", hint: "Uses the system tunnel service if available" },
          { value: "manual", label: "Manual", hint: "I'll set up my own tunnel (ngrok, cloudflared, etc.)" },
        ],
      });

      // ── Step 8d: Save Settings ──

      await settings.setAll({
        enabled: true,
        botAppId,
        botAppPassword,
        tenantId,
        teamId,
        channelId,
        notificationChannelId,
        assistantThreadId: null,
        botPort,
        tunnelMethod,
        ...(graphClientSecret ? { graphClientSecret } : {}),
      });

      terminal.log.success("Teams adapter configured!");
      terminal.log.info("");
      terminal.note(
        `Bot App ID:       ${botAppId}\n` +
        `Tenant:           ${tenantId === "botframework.com" ? "Multi-tenant" : tenantId}\n` +
        `Team ID:          ${teamId}\n` +
        `Channel ID:       ${channelId}\n` +
        `Bot port:         ${botPort}\n` +
        `Notifications:    ${notificationChannelId ?? "Not configured"}\n` +
        `File sharing:     ${graphClientSecret ? "Enabled (Graph API)" : "Disabled"}`,
        "Configuration Summary",
      );

      // ── Step 9: Generate Teams App Package ──

      let appPackagePath: string | null = null;
      try {
        const { generateTeamsAppPackage } = await import("./app-package.js");
        appPackagePath = await generateTeamsAppPackage(botAppId, ctx);
        if (appPackagePath) {
          terminal.log.success(`Teams app package created: ${appPackagePath}`);
        }
      } catch {
        // Non-fatal — user can create manually
      }

      // Tunnel-specific guidance for next steps
      let tunnelStep: string;
      if (tunnelMethod === "devtunnel") {
        tunnelStep =
          `  2. Install the Dev Tunnels plugin:\n` +
          "     openacp plugin install @hahnfeld/devtunnel-provider\n" +
          `     It will automatically tunnel port ${botPort} with a persistent URL.`;
      } else if (tunnelMethod === "builtin") {
        tunnelStep =
          `  2. Tunnel: Will be created automatically on startup via the built-in tunnel service.\n` +
          `     Ensure a tunnel service plugin is installed and running.`;
      } else {
        tunnelStep =
          `  2. Set up a tunnel to expose port ${botPort} (the Bot Framework port):\n` +
          "     Use ngrok, cloudflared, or any tunnel pointing to localhost:" + botPort;
      }

      terminal.log.info("");
      terminal.note(
        "Next steps:\n" +
        "  1. Upload the Teams app package to your team:\n" +
        (appPackagePath
          ? `     File: ${appPackagePath}\n`
          : "     Generate it with: openacp plugin configure @hahnfeld/teams-adapter\n") +
        "     Teams → Apps → Manage your apps → Upload a custom app\n" +
        tunnelStep + "\n" +
        "  3. Set the bot's messaging endpoint in Azure:\n" +
        "     Azure Portal → Bot resource → Configuration → Messaging endpoint\n" +
        "     Example: https://<your-tunnel-url>/api/messages\n" +
        `  4. Note: The bot port (${botPort}) is NOT the same as the OpenACP API port.\n` +
        "     Your tunnel must point to the bot port.\n" +
        "  5. Start OpenACP: openacp start",
        "Next Steps",
      );
    },

    // ─── Configure (post-install changes) ────────────────────────────────

    async configure(ctx: InstallContext) {
      const { terminal, settings } = ctx;
      const current = await settings.getAll();

      const { validateBotCredentials } = await import("./validators.js");

      const choice = await terminal.select({
        message: "What to configure?",
        options: [
          { value: "credentials", label: "Change bot credentials (App ID / Password)" },
          { value: "tenant", label: "Change tenant ID" },
          { value: "team", label: "Change team / channel" },
          { value: "botPort", label: `Change bot port (current: ${(current.botPort as number) ?? DEFAULT_BOT_PORT})` },
          { value: "notifications", label: "Change notification channel" },
          { value: "graph", label: "Configure file sharing (Graph API)" },
          { value: "tunnel", label: `Change tunnel method (current: ${(current.tunnelMethod as string) || "devtunnel"})` },
          { value: "appPackage", label: "Regenerate Teams app package (openacp-bot.zip)" },
          { value: "done", label: "Done" },
        ],
      });

      switch (choice) {
        case "credentials": {
          const appId = await terminal.text({
            message: "Bot App ID:",
            defaultValue: (current.botAppId as string) ?? "",
            validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
          });
          const password = await terminal.password({
            message: "Bot App Password:",
            validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
          });

          const spin = terminal.spinner();
          spin.start("Validating...");
          const existingTenant = (current.tenantId as string) ?? undefined;
          const result = await validateBotCredentials(appId.trim(), password.trim(), existingTenant !== "botframework.com" ? existingTenant : undefined);
          if (result.ok) {
            spin.stop("Credentials validated");
            await settings.set("botAppId", appId.trim());
            await settings.set("botAppPassword", password.trim());
            terminal.log.success("Credentials updated");
          } else {
            spin.fail(result.error);
            const save = await terminal.confirm({ message: "Save anyway?", initialValue: false });
            if (save) {
              await settings.set("botAppId", appId.trim());
              await settings.set("botAppPassword", password.trim());
            }
          }
          break;
        }

        case "tenant": {
          const tid = await terminal.text({
            message: "Tenant ID (GUID, or 'botframework.com' for multi-tenant):",
            defaultValue: (current.tenantId as string) ?? "",
            validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
          });
          await settings.set("tenantId", tid.trim());
          terminal.log.success("Tenant ID updated");
          break;
        }

        case "team": {
          const tid = await terminal.text({
            message: "Team ID:",
            defaultValue: (current.teamId as string) ?? "",
            validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
          });
          const cid = await terminal.text({
            message: "Channel ID:",
            defaultValue: (current.channelId as string) ?? "",
            validate: (v) => (!v.trim() ? "Cannot be empty" : undefined),
          });
          await settings.set("teamId", tid.trim());
          await settings.set("channelId", cid.trim());
          terminal.log.success("Team and channel updated");
          break;
        }

        case "botPort": {
          const portInput = await terminal.text({
            message: `Bot port (Bot Framework HTTP server, default ${DEFAULT_BOT_PORT}):`,
            defaultValue: String((current.botPort as number) ?? DEFAULT_BOT_PORT),
            validate: (v) => {
              const n = Number(v.trim());
              if (isNaN(n) || !Number.isInteger(n) || n < 1 || n > 65535) return "Port must be 1–65535";
              return undefined;
            },
          });
          await settings.set("botPort", Number(portInput.trim()));
          terminal.log.success(`Bot port updated to ${portInput.trim()}`);
          terminal.log.info("Remember to update your tunnel to point to this port.");
          break;
        }

        case "notifications": {
          const nid = await terminal.text({
            message: "Notification channel ID (empty to disable):",
            defaultValue: (current.notificationChannelId as string) ?? "",
          });
          await settings.set("notificationChannelId", nid.trim() || null);
          terminal.log.success(nid.trim() ? "Notification channel updated" : "Notifications disabled");
          break;
        }

        case "graph": {
          const secret = await terminal.password({
            message: "Graph API client secret (empty to disable file sharing):",
          });
          if (secret.trim()) {
            await settings.set("graphClientSecret", secret.trim());
            terminal.log.success("Graph API configured for file sharing");
          } else {
            await settings.delete("graphClientSecret");
            terminal.log.success("File sharing disabled");
          }
          break;
        }

        case "tunnel": {
          const method = await terminal.select({
            message: "How should the bot port be tunneled?",
            options: [
              { value: "devtunnel", label: "@hahnfeld/devtunnel-provider (Recommended)", hint: "Microsoft Dev Tunnels — persistent URLs" },
              { value: "builtin", label: "Built-in tunnel service", hint: "Uses the system tunnel service if available" },
              { value: "manual", label: "Manual", hint: "I'll set up my own tunnel (ngrok, cloudflared, etc.)" },
            ],
          });
          await settings.set("tunnelMethod", method);
          terminal.log.success(`Tunnel method: ${method === "builtin" ? "built-in (auto-create)" : method === "devtunnel" ? "devtunnel plugin" : "manual"}`);
          if (method === "devtunnel") {
            terminal.log.info("Install with: openacp plugin install @hahnfeld/devtunnel-provider");
          }
          break;
        }

        case "appPackage": {
          const appId = (current.botAppId as string) ?? "";
          if (!appId) {
            terminal.log.error("Bot App ID is not configured. Set credentials first.");
            break;
          }
          try {
            const { generateTeamsAppPackage } = await import("./app-package.js");
            const path = await generateTeamsAppPackage(appId, ctx);
            if (path) {
              terminal.log.success(`Teams app package created: ${path}`);
              terminal.log.info("Upload it in Teams → Apps → Manage your apps → Upload a custom app");
            } else {
              terminal.log.error("Failed to generate app package");
            }
          } catch (err) {
            terminal.log.error(`App package generation failed: ${err instanceof Error ? err.message : String(err)}`);
          }
          break;
        }

        case "done":
          break;
      }
    },

    // ─── Uninstall ───────────────────────────────────────────────────────

    async uninstall(ctx: InstallContext, opts: { purge: boolean }) {
      if (opts.purge) {
        await ctx.settings.clear();
        ctx.terminal.log.success("Teams adapter settings cleared");
      }
      ctx.terminal.note(
        "Don't forget to:\n" +
        "  1. Remove the bot from your Teams team (if no longer needed)\n" +
        "  2. Delete the Azure Bot resource (if no longer needed)\n" +
        "     Azure Portal → Bot resource → Delete",
        "Cleanup Reminder",
      );
    },

    // ─── Runtime Setup ───────────────────────────────────────────────────

    async setup(ctx) {
      const rawPort = (ctx.pluginConfig as Record<string, unknown>).botPort;
      const botPort = typeof rawPort === "number" ? rawPort : (typeof rawPort === "string" ? Number(rawPort) : DEFAULT_BOT_PORT) || DEFAULT_BOT_PORT;

      ctx.registerEditableFields([
        { key: "enabled", displayName: "Enabled", type: "toggle", scope: "safe", hotReload: false },
        { key: "botAppId", displayName: "Bot App ID", type: "string", scope: "sensitive", hotReload: false },
        { key: "tenantId", displayName: "Tenant ID", type: "string", scope: "safe", hotReload: false },
        { key: "teamId", displayName: "Team ID", type: "string", scope: "safe", hotReload: false },
        { key: "channelId", displayName: "Channel ID", type: "string", scope: "safe", hotReload: false },
        { key: "botPort", displayName: "Bot Port", type: "number", scope: "safe", hotReload: false },
        { key: "tunnelMethod", displayName: "Tunnel method", type: "string", scope: "safe", hotReload: false },
      ]);

      const config = ctx.pluginConfig as Record<string, unknown>;
      if (!config.enabled || !config.botAppId) {
        ctx.log.info("Teams adapter disabled (missing enabled or botAppId)");
        return;
      }

      if (adapter) {
        ctx.log.warn("Teams adapter setup() called again — skipping (already running)");
        return;
      }
      adapter = new TeamsAdapter(ctx.core as OpenACPCore, config as unknown as TeamsChannelConfig);
      ctx.registerService("adapter:teams", adapter);
      ctx.log.info("Teams adapter registered");

      // Only create a tunnel if the user opted into the built-in tunnel service during install.
      // Most users should use @hahnfeld/devtunnel-provider instead (persistent URLs, auto-managed).
      const tunnelMethod = ((ctx.pluginConfig as Record<string, unknown>).tunnelMethod as string) || "devtunnel";

      if (tunnelMethod === "builtin") {
        // Note: addTunnel() exists on the concrete TunnelService class but is not on
        // the exported TunnelServiceInterface type — use runtime typeof check.
        const tunnelSvc = ctx.getService("tunnel") as Record<string, unknown> | undefined;
        if (tunnelSvc && typeof tunnelSvc.addTunnel === "function") {
          try {
            const entry = await (tunnelSvc.addTunnel as (port: number, opts?: { label?: string }) => Promise<{ publicUrl?: string }>)(
              botPort, { label: "teams-bot" },
            );
            if (entry?.publicUrl) {
              ctx.log.info(`Teams bot tunnel active — messaging endpoint: ${entry.publicUrl}/api/messages`);
              ctx.log.info("Set this URL as the messaging endpoint in Azure Portal → Bot resource → Configuration");
            }
          } catch (err) {
            ctx.log.warn(`Could not create tunnel for bot port ${botPort}: ${(err as Error).message}`);
            ctx.log.info(`Tunnel your bot manually: <tunnel-url> → localhost:${botPort}`);
            ctx.log.info("Tip: Install @hahnfeld/devtunnel-provider for Microsoft Dev Tunnels integration");
          }
        } else {
          ctx.log.warn(`Auto-tunnel enabled but no tunnel service available — install a tunnel plugin`);
          ctx.log.info("Recommended: openacp plugin install @hahnfeld/devtunnel-provider");
        }
      } else {
        ctx.log.info(`Bot listening on port ${botPort} — tunnel method: ${tunnelMethod}`);
      }
    },

    async teardown() {
      if (adapter) {
        await adapter.stop();
        adapter = null;
      }
    },
  };
}
