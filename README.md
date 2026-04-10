# @openacp/teams-adapter

Microsoft Teams adapter plugin for [OpenACP](https://github.com/Open-ACP/OpenACP) — Adaptive Cards, slash commands, streaming.

## Features

- **Adaptive Cards** — Rich tool card rendering with progress indicators and action buttons
- **Slash Commands** — Full command suite: `/new`, `/cancel`, `/agents`, `/menu`, and more
- **Streaming** — Real-time text updates via Teams message editing
- **Threads** — Session threads within Teams channels
- **Permissions** — Allow/Deny/Always Allow via Adaptive Card buttons
- **Output Modes** — Low/Medium/High detail levels
- **File Sharing** — Upload and share files via Microsoft Graph / OneDrive
- **Interactive Install Wizard** — Guided setup with credential validation and auto-discovery

## Prerequisites

- [OpenACP CLI](https://github.com/Open-ACP/OpenACP) `>= 2026.0.0`
- Node.js 18+
- An **Azure Bot registration** with the Microsoft Teams channel enabled
- A Microsoft 365 tenant with a Teams team and channel

## Installation

### Option A: OpenACP plugin install (recommended)

```bash
openacp plugin install @openacp/teams-adapter
```

This launches an interactive wizard that walks you through Azure Bot setup, credential validation, and team/channel selection.

### Option B: Manual npm install

```bash
npm install @openacp/teams-adapter
# or
pnpm add @openacp/teams-adapter
```

## Azure Bot Setup

Before configuring the adapter you need an Azure Bot registration. If you don't have one yet:

1. Go to the [Azure Bot creation page](https://portal.azure.com/#create/Microsoft.AzureBot)
2. Fill in:
   - **Bot handle**: any unique name (e.g. `openacp-bot`)
   - **Pricing tier**: Free (F0) for testing
   - **App type**: "Single Tenant" for enterprise use, "Multi Tenant" for public bots
   - **Creation type**: "Create new Microsoft App ID"
3. Click **Create** and wait for deployment
4. Go to the Bot resource > **Settings** > **Configuration**
   - Copy the **Microsoft App ID** — this is your `botAppId`
5. Click **Manage Password** > **New client secret**
   - Copy the secret value immediately (shown only once) — this is your `botAppPassword`
6. Under **Channels**, add the **Microsoft Teams** channel

For full details see the [Azure Bot Service docs](https://learn.microsoft.com/en-us/azure/bot-service/bot-service-quickstart-registration).

## Configuration

### Interactive wizard

If you installed via `openacp plugin install`, the wizard runs automatically. To re-run it later:

```bash
openacp plugin configure @openacp/teams-adapter
```

The wizard guides you through:
1. Bot credential entry and real-time validation
2. Tenant type selection (single-tenant enterprise vs. multi-tenant)
3. Team and channel selection — auto-discovered via Graph API, or paste a Teams channel link
4. Optional notification channel for session completions and errors
5. Optional Graph API file sharing (OneDrive)

### Manual configuration

Add the following to your `openacp.yaml`:

```yaml
channels:
  teams:
    enabled: true
    botAppId: "${TEAMS_BOT_APP_ID}"
    botAppPassword: "${TEAMS_BOT_APP_PASSWORD}"
    tenantId: "${TEAMS_TENANT_ID}"
    teamId: "${TEAMS_TEAM_ID}"
    channelId: "${TEAMS_CHANNEL_ID}"
    notificationChannelId: "${TEAMS_NOTIFICATION_CHANNEL_ID}"  # optional
    assistantThreadId: null  # auto-set after first run
    graphClientSecret: "${TEAMS_GRAPH_CLIENT_SECRET}"           # optional, for file sharing
```

### Configuration reference

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `enabled` | `boolean` | Yes | Enable the Teams adapter |
| `botAppId` | `string` | Yes | Azure AD App ID for the bot |
| `botAppPassword` | `string` | Yes | Bot client secret |
| `tenantId` | `string` | Yes | Microsoft tenant ID (GUID), or `botframework.com` for multi-tenant |
| `teamId` | `string` | Yes | Default team ID (groupId GUID) |
| `channelId` | `string` | Yes | Primary channel for sessions (e.g. `19:abc@thread.tacv2`) |
| `notificationChannelId` | `string \| null` | No | Separate channel for notifications |
| `assistantThreadId` | `string \| null` | No | Thread for the assistant (auto-populated) |
| `graphClientSecret` | `string` | No | Azure AD client secret for Graph API file sharing |

### Finding your Team and Channel IDs

**From a channel link (easiest):**
1. Open Microsoft Teams
2. Right-click the channel name > **Get link to channel**
3. The link contains `groupId` (Team ID) and the channel path

**From the Azure/Graph API:**
- Team ID = the `groupId` parameter from the Teams URL
- Channel ID = the encoded string like `19:xxx@thread.tacv2`

### Messaging endpoint

After installation, set the bot's messaging endpoint in Azure:

1. Azure Portal > Bot resource > **Configuration** > **Messaging endpoint**
2. Set it to: `https://your-server.com/api/messages`
3. Make sure your server is publicly reachable (use a tunnel like ngrok for local development)

## Slash Commands

| Command | Description |
|---------|-------------|
| `/new [agent]` | Create a new agent session |
| `/newchat` | New chat, same agent & workspace |
| `/cancel` | Cancel the current session |
| `/status` | Show session or global status |
| `/sessions` | List all sessions |
| `/agents` | List available agents |
| `/install <name>` | Install an agent by name |
| `/menu` | Show the action menu |
| `/help` | Show help |
| `/outputmode low\|medium\|high` | Set output detail level |
| `/bypass` | Auto-approve permissions |
| `/doctor` | Run system diagnostics |
| `/handoff` | Generate terminal resume command |
| `/restart` | Restart OpenACP |
| `/respawn` | Restart the assistant session |
| `/update` | Update to latest version |
| `/settings` | Show configuration settings |
| `/integrate` | Manage agent integrations |
| `/clear` | Reset the assistant session |
| `/tts [on\|off]` | Toggle Text to Speech |
| `/mode` | Switch session mode |
| `/model` | Switch AI model |
| `/thought` | Adjust thinking level |

## Uninstalling

```bash
openacp plugin uninstall @openacp/teams-adapter --purge
```

The `--purge` flag removes all saved settings. After uninstalling, you may also want to:
1. Remove the bot from your Teams team
2. Delete the Azure Bot resource in the Azure Portal

## Development

```bash
# Install dependencies
pnpm install

# Build
pnpm build

# Watch mode
pnpm dev

# Run tests
pnpm test
```

## Architecture

```
src/
├── index.ts             # Plugin entry point & public exports
├── plugin.ts            # Plugin factory (install wizard, configure, setup/teardown)
├── adapter.ts           # TeamsAdapter — extends MessagingAdapter
├── renderer.ts          # TeamsRenderer (Adaptive Card rendering)
├── activity.ts          # ActivityTracker (tool card state, streaming)
├── formatting.ts        # Tool card formatting, usage cards, citations
├── draft-manager.ts     # Message draft handling
├── permissions.ts       # PermissionHandler (Adaptive Card buttons)
├── graph.ts             # GraphFileClient (OneDrive file sharing)
├── media.ts             # File download/upload utilities
├── conversation-store.ts # Conversation reference storage
├── send-utils.ts        # Message sending helpers
├── task-modules.ts      # Task module dialogs (new session, settings)
├── assistant.ts         # Assistant session spawning
├── validators.ts        # Credential & tenant validation, Teams link parsing
├── types.ts             # TeamsChannelConfig, TeamsPlatformData
└── commands/
    ├── index.ts          # Command router + SLASH_COMMANDS registry
    ├── new-session.ts    # /new, /newchat
    ├── session.ts        # /cancel, /status, /sessions, /handoff
    ├── admin.ts          # /bypass, /tts, /restart, /respawn, /update, /outputmode
    ├── menu.ts           # /menu, /help, /clear
    ├── agents.ts         # /agents, /install
    ├── doctor.ts         # /doctor
    ├── integrate.ts      # /integrate
    └── settings.ts       # /settings
```

## Tech Stack

- [`@microsoft/teams.apps`](https://www.npmjs.com/package/@microsoft/teams.apps) — App class, server hosting, activity routing
- [`@microsoft/teams.botbuilder`](https://www.npmjs.com/package/@microsoft/teams.botbuilder) — Bot Framework adapter plugin
- [`@microsoft/agents-hosting`](https://www.npmjs.com/package/@microsoft/agents-hosting) — Express server hosting
- [`adaptivecards-templating`](https://www.npmjs.com/package/adaptivecards-templating) — Adaptive Card templating
- [`@openacp/plugin-sdk`](https://github.com/Open-ACP/OpenACP) — OpenACP plugin interface

## License

MIT
