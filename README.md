# @openacp/teams-adapter

Microsoft Teams adapter plugin for [OpenACP](https://github.com/Open-ACP/OpenACP) — Adaptive Cards, slash commands, streaming.

## Features

- **Adaptive Cards** — Rich tool card rendering with progress indicators, action buttons
- **Slash Commands** — Full command suite: `/new`, `/cancel`, `/agents`, `/menu`, etc.
- **Streaming** — Real-time text updates via Teams message editing
- **Threads** — Session threads within Teams channels
- **Permissions** — Allow/Deny/Always Allow Adaptive Card buttons
- **Output Modes** — Low/Medium/High detail levels

## Installation

```bash
npm install @openacp/teams-adapter
```

Or add to your `openacp.yaml`:

```yaml
channels:
  teams:
    enabled: true
    botAppId: "${TEAMS_BOT_APP_ID}"
    botAppPassword: "${TEAMS_BOT_APP_PASSWORD}"
    tenantId: "${TEAMS_TENANT_ID}"
    teamId: "${TEAMS_TEAM_ID}"
    channelId: "${TEAMS_CHANNEL_ID}"
    notificationChannelId: "${TEAMS_NOTIFICATION_CHANNEL_ID}"
    assistantThreadId: null  # Set after first run
```

## Configuration

| Field | Type | Description |
|-------|------|-------------|
| `enabled` | `boolean` | Enable the Teams adapter |
| `botAppId` | `string` | Azure AD App ID for the bot |
| `botAppPassword` | `string` | App password |
| `tenantId` | `string` | Microsoft tenant ID |
| `teamId` | `string` | Default team ID |
| `channelId` | `string` | Primary channel for sessions |
| `notificationChannelId` | `string \| null` | Channel for notifications |
| `assistantThreadId` | `string \| null` | Thread for the assistant |

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
| `/update` | Update to latest version |
| `/settings` | Show configuration settings |
| `/integrate` | Manage agent integrations |
| `/clear` | Reset the assistant session |
| `/tts [on\|off]` | Toggle Text to Speech |

## Development

```bash
# Install dependencies
pnpm install

# Build
pnpm build

# Watch mode
pnpm dev

# Test
pnpm test
```

## Architecture

```
teams-adapter/
├── src/
│   ├── adapter.ts        # TeamsAdapter extends MessagingAdapter
│   ├── renderer.ts        # TeamsRenderer (Adaptive Cards)
│   ├── activity.ts       # ActivityTracker (tool card state, streaming)
│   ├── formatting.ts     # Tool card formatting, usage, permissions
│   ├── draft-manager.ts  # Message draft handling
│   ├── permissions.ts    # PermissionHandler (Adaptive Cards)
│   ├── types.ts          # TeamsChannelConfig
│   ├── commands/
│   │   ├── index.ts      # Command router + SLASH_COMMANDS
│   │   ├── new-session.ts
│   │   ├── session.ts
│   │   ├── admin.ts
│   │   ├── menu.ts
│   │   ├── agents.ts
│   │   ├── doctor.ts
│   │   ├── integrate.ts
│   │   └── settings.ts
│   └── index.ts
└── index.ts              # Plugin entry point
```

## Tech Stack

- `@microsoft/teams.apps` — App class, server hosting, activity routing
- `@microsoft/teams.botbuilder` — Adapter plugin integrating Bot Framework
- `@microsoft/teams.cards` — Adaptive Card builders and typings
- `@microsoft/agents-hosting` — Express server hosting
- `adaptivecards-templating` — Adaptive Card templating

## License

MIT