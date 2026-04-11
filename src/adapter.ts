import { App } from "@microsoft/teams.apps";
import { BotBuilderPlugin } from "@microsoft/teams.botbuilder";
import { CardFactory, MemoryStorage } from "@microsoft/agents-hosting";
import type { TurnContext } from "@microsoft/agents-hosting";
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
} from "botbuilder";
import { PasswordServiceClientCredentialFactory } from "botframework-connector";

import type { InvokeResponse } from "@microsoft/teams.api";
import type {
  OutgoingMessage,
  PermissionRequest,
  NotificationMessage,
  PlanEntry,
  Attachment,
  OpenACPCore,
  Session,
  DisplayVerbosity,
  AdapterCapabilities,
  IRenderer,
  MessagingAdapterConfig,
  FileServiceInterface,
  CommandResponse,
} from "@openacp/plugin-sdk";
import { log, MessagingAdapter, SendQueue } from "@openacp/plugin-sdk";
import type { CommandRegistry } from "@openacp/plugin-sdk";
import { TeamsRenderer } from "./renderer.js";
import type { TeamsChannelConfig } from "./types.js";
import { DEFAULT_BOT_PORT } from "./types.js";
import { TeamsDraftManager } from "./draft-manager.js";
import { PermissionHandler } from "./permissions.js";
import { handleCommand, setupCardActionCallbacks, SLASH_COMMANDS } from "./commands/index.js";
import { spawnAssistant } from "./assistant.js";
import { downloadTeamsFile, isAttachmentTooLarge, buildFileAttachmentCard, uploadFileViaGraph } from "./media.js";
import { GraphFileClient } from "./graph.js";
import { ConversationStore } from "./conversation-store.js";
import { sendText, sendCard, sendActivity } from "./send-utils.js";
import { renderUsageCard, renderToolCallCard, renderPlanCard, buildCitationEntities } from "./formatting.js";
import { buildNewSessionDialog, buildSettingsDialog, buildDialogMessage } from "./task-modules.js";
import type { ToolCallMeta } from "@openacp/plugin-sdk";
import type { OutputMode } from "./activity.js";

/** Max retry attempts for transient Teams API failures */
const MAX_RETRIES = 3;
/** Base delay (ms) for exponential backoff */
const BASE_RETRY_DELAY = 1000;

export class TeamsAdapter extends MessagingAdapter {
  readonly name = "teams";
  readonly renderer: IRenderer = new TeamsRenderer();
  readonly capabilities: AdapterCapabilities = {
    streaming: true,
    richFormatting: true,
    threads: true,
    reactions: false,
    fileUpload: true,
    voice: false,
  };

  readonly core: OpenACPCore;
  private app: App;
  private teamsConfig: TeamsChannelConfig;
  private sendQueue: SendQueue;
  private draftManager: TeamsDraftManager;
  private permissionHandler!: PermissionHandler;

  private notificationChannelId?: string;
  private assistantSession: Session | null = null;
  private assistantInitializing = false;
  private fileService: FileServiceInterface;
  private graphClient?: GraphFileClient;
  private conversationStore: ConversationStore;

  /**
   * Per-session TurnContext references, set during inbound message handling.
   * Handler overrides read from this map during sendMessage dispatch.
   */
  private _sessionContexts = new Map<string, { context: TurnContext; isAssistant: boolean; threadId?: string }>();
  private _sessionOutputModes = new Map<string, OutputMode>();

  /**
   * Per-session serial dispatch queues — matches Telegram's _dispatchQueues pattern.
   * SessionBridge fires sendMessage() as fire-and-forget, so multiple events can arrive
   * concurrently. Without serialization, fast handlers overtake slow ones, causing
   * out-of-order delivery. This queue ensures events are processed in arrival order.
   *
   * Entries are replaced with Promise.resolve() once their chain settles, preventing
   * unbounded closure growth for long-lived sessions.
   */
  private _dispatchQueues = new Map<string, Promise<void>>();

  /** Track processed activity IDs to handle Teams 15-second retry deduplication */
  private _processedActivities = new Map<string, number>();
  private _processedCleanupTimer?: ReturnType<typeof setInterval>;

  /** Messages buffered during assistant initialization — replayed once ready. Capped to prevent unbounded growth. */
  private static readonly MAX_INIT_BUFFER = 50;
  private _assistantInitBuffer: Array<{ sessionId: string; content: OutgoingMessage }> = [];

  /** Bot token cache for proactive messaging via connector REST API */
  private _botTokenCache?: { token: string; expiresAt: number };

  constructor(core: OpenACPCore, config: TeamsChannelConfig) {
    super(
      { configManager: core.configManager },
      // Teams measures message size in bytes (100KB limit, 80KB safe threshold).
      // Use 12000 chars as the split limit — safe for multi-byte content (CJK, emoji)
      // where each char can be 3-4 bytes, plus activity envelope overhead.
      { ...config as unknown as Record<string, unknown>, maxMessageLength: 12000, enabled: config.enabled ?? true } as MessagingAdapterConfig,
    );
    this.core = core;
    this.teamsConfig = config;
    this.sendQueue = new SendQueue({ minInterval: 1000 });
    this.draftManager = new TeamsDraftManager(this.sendQueue);
    this.fileService = core.fileService;

    // Persistent conversation reference store for proactive messaging
    const storageDir = (core.configManager as any).instanceRoot ?? process.cwd();
    this.conversationStore = new ConversationStore(storageDir);

    // Initialize Graph API client for file operations (optional)
    if (config.graphClientSecret && config.tenantId && config.botAppId) {
      this.graphClient = new GraphFileClient(
        config.tenantId,
        config.botAppId,
        config.graphClientSecret,
      );
      log.info("[TeamsAdapter] Graph API file client initialized");
    }

    const isSingleTenant = config.tenantId && config.tenantId !== "botframework.com";

    // Custom credential factory that accepts both the bot's app ID and the
    // Bot Framework audience URI. Azure Bot Service sends single-tenant tokens
    // with aud=https://api.botframework.com, but the SDK's default factory only
    // accepts the bot's own app ID — causing "Invalid AppId" errors.
    const credFactory = new PasswordServiceClientCredentialFactory(
      config.botAppId,
      config.botAppPassword,
      isSingleTenant ? config.tenantId : "",
    );
    const origIsValidAppId = credFactory.isValidAppId.bind(credFactory);
    credFactory.isValidAppId = async (appId: string): Promise<boolean> => {
      if (appId === "https://api.botframework.com") return true;
      return origIsValidAppId(appId);
    };

    const botAuth = new ConfigurationBotFrameworkAuthentication({}, credFactory);
    const cloudAdapter = new CloudAdapter(botAuth);
    const botBuilderPlugin = new BotBuilderPlugin({ adapter: cloudAdapter });
    this.app = new App({
      clientId: config.botAppId,
      clientSecret: config.botAppPassword,
      tenantId: isSingleTenant ? config.tenantId : undefined,
      port: config.botPort ?? DEFAULT_BOT_PORT,
      skipAuth: true, // App-level JWT validation skipped; CloudAdapter handles auth
      storage: new MemoryStorage(),
      plugins: [botBuilderPlugin],
    } as any);
  }

  // ─── start ────────────────────────────────────────────────────────────────

  async start(): Promise<void> {
    log.info("[TeamsAdapter] Starting...");

    try {
      this.notificationChannelId = this.teamsConfig.notificationChannelId ?? undefined;

      this.permissionHandler = new PermissionHandler(
        (sessionId) => this.core.sessionManager.getSession(sessionId),
        (notification) => this.sendNotification(notification),
      );

      this.setupMessageHandler();
      this.setupCardActionHandler();
      this.setupTaskModuleHandlers();
      this.setupReactionHandler();

      // Periodic cleanup of processed activity IDs (deduplication cache)
      this._processedCleanupTimer = setInterval(() => {
        const cutoff = Date.now() - 60_000; // 60s retention
        for (const [id, ts] of this._processedActivities) {
          if (ts < cutoff) this._processedActivities.delete(id);
        }
      }, 30_000);
      if (this._processedCleanupTimer.unref) this._processedCleanupTimer.unref();

      await this.app.start();
      const botPort = this.teamsConfig.botPort ?? DEFAULT_BOT_PORT;
      log.info(`[TeamsAdapter] Bot Framework server listening on port ${botPort}`);

      // Spawn assistant session if configured (non-blocking — matches Telegram's pattern)
      if (this.teamsConfig.assistantThreadId) {
        this.setupAssistant().catch((err) => {
          log.error({ err }, "[TeamsAdapter] Assistant setup failed (non-blocking)");
        });
      }

      log.info("[TeamsAdapter] Initialization complete");
    } catch (err) {
      log.error({ err }, "[TeamsAdapter] Initialization failed");
      throw err;
    }
  }

  // ─── stop ─────────────────────────────────────────────────────────────────

  async stop(): Promise<void> {
    // Cancel deduplication cleanup timer
    if (this._processedCleanupTimer) {
      clearInterval(this._processedCleanupTimer);
      this._processedCleanupTimer = undefined;
    }

    if (this.assistantSession) {
      try {
        await this.assistantSession.destroy();
      } catch (err) {
        log.warn({ err }, "[TeamsAdapter] Failed to destroy assistant session");
      }
      this.assistantSession = null;
    }

    this._sessionContexts.clear();
    this._sessionOutputModes.clear();
    this._dispatchQueues.clear();
    this._processedActivities.clear();
    this.sendQueue.clear();
    this.conversationStore.destroy();
    this.permissionHandler.dispose();

    await this.app.stop();
    log.info("[TeamsAdapter] Stopped");
  }

  // ─── Message handler ──────────────────────────────────────────────────────

  private setupMessageHandler(): void {
    // File consent: accept
    this.app.on("file.consent.accept", async (context: any) => {
      try {
        const uploadInfo = context.activity.value?.uploadInfo;
        const fileName = uploadInfo?.name ?? "file";
        log.info({ fileName }, "[TeamsAdapter] File consent accepted");
        await sendText(context, `✅ File consent accepted: ${fileName}`);
        return { status: 200 };
      } catch (err) {
        log.error({ err }, "[TeamsAdapter] file.consent.accept error");
        return { status: 500 };
      }
    });

    // File consent: decline
    this.app.on("file.consent.decline", async (context: any) => {
      try {
        const uploadInfo = context.activity.value?.uploadInfo;
        const fileName = uploadInfo?.name ?? "file";
        log.info({ fileName }, "[TeamsAdapter] File consent declined");
        await sendText(context, `❌ File consent declined: ${fileName}`);
        return { status: 200 };
      } catch (err) {
        log.error({ err }, "[TeamsAdapter] file.consent.decline error");
        return { status: 500 };
      }
    });

    this.app.on("message", async (context: any) => {
      // Action.Submit cards send activity.value with the flat data object (no text).
      // Intercept these before normal text processing.
      const submitValue = context.activity.value as Record<string, unknown> | undefined;
      if (submitValue && !context.activity.text) {
        await this.handleSubmitAction(context, submitValue);
        return;
      }

      const rawActivityText = context.activity.text ?? "";
      // Teams may prepend the bot @mention (e.g., "<at>BotName</at> /new") — strip it
      const text = rawActivityText.replace(/<at[^>]*>.*?<\/at>/gi, "").trim();
      log.info({ rawText: rawActivityText.slice(0, 100), cleanText: text.slice(0, 100), activityType: context.activity.type }, "[TeamsAdapter] Incoming activity");

      const userId = context.activity.from?.id ?? "unknown";
      // Use conversation.id as the thread discriminator — NOT activity.channelId
      // which is always "msteams" for Teams. Conversation ID uniquely identifies
      // the 1:1, group chat, or channel thread the message came from.
      const conversationId = String(context.activity.conversation?.id ?? "unknown");
      const threadId = conversationId;
      const activityId = context.activity.id as string | undefined;

      try {
        // Idempotency: Teams retries if bot takes >15s — skip duplicate activity IDs
        if (activityId && this._processedActivities.has(activityId)) {
          log.debug({ activityId }, "[TeamsAdapter] Duplicate activity, skipping");
          return;
        }
        if (activityId) this._processedActivities.set(activityId, Date.now());

        log.debug({ threadId, userId, text: text.slice(0, 50) }, "[TeamsAdapter] message received");

        if (!text && !context.activity.attachments?.length) return;

        // Persist conversation reference for proactive messaging (survives restarts).
        // Validate serviceUrl against known Bot Framework endpoints to prevent
        // SSRF via spoofed serviceUrl redirecting bot tokens to an attacker.
        if (context.activity.conversation?.id && context.activity.serviceUrl) {
          const serviceUrl = context.activity.serviceUrl as string;
          if (TeamsAdapter.isValidServiceUrl(serviceUrl)) {
            this.conversationStore.upsert({
              conversationId: context.activity.conversation.id,
              serviceUrl,
              tenantId: context.activity.conversation.tenantId ?? this.teamsConfig.tenantId,
              channelId: context.activity.channelId,
              botId: context.activity.recipient?.id ?? this.teamsConfig.botAppId,
              botName: context.activity.recipient?.name ?? "OpenACP",
              updatedAt: Date.now(),
            });
          } else {
            log.warn({ serviceUrl: serviceUrl.slice(0, 80) }, "[TeamsAdapter] Rejected untrusted serviceUrl");
          }
        }

        const existingSession = this.core.sessionManager.getSessionByThread("teams", threadId);
        const sessionId = existingSession?.id ?? "unknown";

        // Always store context under threadId — this is the stable key.
        // For new sessions, core.handleMessage creates the session and assigns threadId,
        // then sendMessage is called with the new sessionId. We also store under sessionId
        // when known, and under threadId as a universal fallback.
        const isAssistant = this.assistantSession != null && sessionId === this.assistantSession?.id;
        this._sessionContexts.set(threadId, { context, isAssistant, threadId });
        if (sessionId !== "unknown") {
          this._sessionContexts.set(sessionId, { context, isAssistant, threadId });
        }

        // Route slash commands — local handler first (handles /new, /help, etc.
        // with Teams-specific UX), then fall back to core command registry.
        if (text.startsWith("/")) {
          // Always try local command handler first — it has Teams-specific implementations
          const handled = await handleCommand(context, this, userId, sessionId !== "unknown" ? sessionId : null);
          if (handled) return;

          // Fall back to core command registry for commands not handled locally
          const registry = this.getCommandRegistry();
          if (registry) {
            const rawCommand = text.split(" ")[0].slice(1).toLowerCase();
            const def = registry.get(rawCommand);
            if (def) {
              try {
                const response = await registry.execute(text, {
                  raw: "",
                  sessionId: sessionId !== "unknown" ? sessionId : null,
                  channelId: "teams",
                  userId,
                  reply: async (content: string) => {
                    if (typeof content === "string") {
                      await sendText(context, content);
                    }
                  },
                });
                if (response.type !== "silent") {
                  await this.renderCommandResponse(response, context);
                }
              } catch (err) {
                log.error({ err, command: rawCommand }, "[TeamsAdapter] Command failed");
                await sendText(context, `❌ Command failed: ${err instanceof Error ? err.message : String(err)}`);
              }
              return;
            }
          }
          return;
        }

        // Process attachments only when a session already exists — saving files
        // requires a valid sessionId for the file service to scope storage correctly.
        // First messages (sessionId === "unknown") skip attachment processing; the
        // auto-session-creation path below handles the text, and the user can resend
        // the attachment once the session is established.
        const attachments: Attachment[] = [];
        if (sessionId !== "unknown" && context.activity.attachments?.length) {
          for (const att of context.activity.attachments) {
            try {
              const buffer = await downloadTeamsFile(
                att.contentUrl ?? "",
                att.name ?? "attachment",
                this.graphClient,
              );
              if (buffer) {
                const saved = await this.fileService.saveFile(
                  sessionId,
                  att.name ?? "attachment",
                  buffer,
                  att.contentType ?? "application/octet-stream",
                );
                if (saved) attachments.push(saved);
              }
            } catch (err) {
              log.warn({ err, name: att.name }, "[TeamsAdapter] Attachment download failed");
            }
          }
        }

        let messageText = text;
        if (!messageText && attachments.length > 0) {
          messageText = attachments.map((a) => `[Attachment: ${a.fileName}]`).join("\n");
        }

        if (!messageText && attachments.length === 0) return;

        // Route to assistant thread
        if (
          this.teamsConfig.assistantThreadId &&
          threadId === this.teamsConfig.assistantThreadId
        ) {
          if (this.assistantSession && messageText) {
            try {
              await this.assistantSession.enqueuePrompt(
                messageText,
                attachments.length > 0 ? attachments : undefined,
              );
            } catch (err) {
              log.error({ err, sessionId: this.assistantSession.id }, "[TeamsAdapter] assistant enqueuePrompt failed");
            }
          }
          return;
        }

        if (sessionId !== "unknown") {
          // Drain pending dispatches and reset draft state for new prompt
          const pendingDispatch = this._dispatchQueues.get(sessionId);
          if (pendingDispatch) await pendingDispatch;
          this.draftManager.cleanup(sessionId);
        }

        // Show typing indicator while the agent processes the message
        this.sendTyping(context);

        const existingSessionBeforeSend = this.core.sessionManager.getSessionByThread("teams", threadId);
        if (!existingSessionBeforeSend) {
          const defaultAgent = (this.core.configManager.get() as Record<string, unknown>)?.defaultAgent as string ?? "claude";
          log.info({ threadId, text: messageText.slice(0, 50), defaultAgent }, "[TeamsAdapter] No session — auto-creating via /new");
          // Auto-create a session for first-time messages
          const registry = this.getCommandRegistry();
          if (registry) {
            try {
              const response = await registry.execute(`/new ${defaultAgent}`, {
                raw: messageText,
                sessionId: null,
                channelId: "teams",
                userId,
                reply: async (content: string) => {
                  if (typeof content === "string") {
                    await sendText(context, content);
                  }
                },
              });
              if (response.type !== "silent") {
                await this.renderCommandResponse(response, context);
              }
              // After session creation, route the original message to the new session
              const newSession = this.core.sessionManager.getSessionByThread("teams", threadId);
              if (newSession && messageText) {
                this._sessionContexts.set(newSession.id, { context, isAssistant: false });
                await this.core.handleMessage({
                  channelId: "teams",
                  threadId,
                  userId,
                  text: messageText,
                  ...(attachments.length > 0 ? { attachments } : {}),
                });
              }
            } catch (err) {
              log.error({ err }, "[TeamsAdapter] Auto-create session failed");
              await sendText(context, "👋 Send /new to start a session.");
            }
            return;
          }
          await sendText(context, "👋 Send /new to start a session.");
          return;
        }

        await this.core.handleMessage({
          channelId: "teams",
          threadId,
          userId,
          text: messageText,
          ...(attachments.length > 0 ? { attachments } : {}),
        });
      } catch (err) {
        log.error({ err, threadId, userId }, "[TeamsAdapter] message handler error");
        try {
          await sendText(context, "❌ Failed to process message. Please try again.");
        } catch { /* best effort */ }
      }
    });
  }

  // ─── Card action handler ────────────────────────────────────────────────

  /**
   * Handle Action.Submit payloads from Adaptive Cards.
   *
   * Action.Submit (v1.2 compatible) sends the card's data directly as
   * activity.value with no text. This handles permission responses, command
   * buttons, and output mode changes — all of which embed a `verb` field.
   */
  private async handleSubmitAction(context: any, data: Record<string, unknown>): Promise<void> {
    try {
      // Task Module dialog triggers: Action.Submit with msteams.type = "task/fetch"
      // sends dialogId but no verb. Handle these by sending the dialog card inline.
      const dialogId = data.dialogId as string | undefined;
      if (dialogId && !data.verb) {
        await this.handleInlineDialog(context, dialogId, data);
        return;
      }

      const verb = data.verb as string | undefined;
      if (!verb) return;

      await this.dispatchCardVerb(context, verb, data);
    } catch (err) {
      log.error({ err }, "[TeamsAdapter] handleSubmitAction error");
    }
  }

  /**
   * Shared card action dispatch — handles verbs from both Action.Submit and
   * Action.Execute (invoke) paths. Eliminates duplication between
   * handleSubmitAction and cardActionHandler.
   */
  private async dispatchCardVerb(context: any, verb: string, data: Record<string, unknown>): Promise<void> {
    // Permission response: data has verb + sessionId + callbackKey + requestId
    if (data.sessionId && data.callbackKey && data.requestId) {
      const handled = await this.permissionHandler.handleCardAction(
        context,
        verb,
        data.sessionId as string,
        data.callbackKey as string,
        data.requestId as string,
      );
      if (handled) return;
    }

    // Output mode change: verb = "om:<sessionId>:<mode>"
    if (verb.startsWith("om:")) {
      const parts = verb.split(":");
      if (parts.length === 3) {
        const [, sessionId, mode] = parts;
        if (mode === "low" || mode === "medium" || mode === "high") {
          const session = this.core.sessionManager.getSession(sessionId);
          if (!session) {
            log.warn({ sessionId }, "[TeamsAdapter] output mode change: session not found");
            return;
          }
          // Verify the action came from the session's conversation
          const conversationId = context.activity.conversation?.id;
          const sessionThread = session.threadIds.get("teams");
          if (sessionThread && conversationId && sessionThread !== conversationId) {
            log.warn({ sessionId, conversationId, sessionThread }, "[TeamsAdapter] om: conversation mismatch");
            return;
          }
          this._sessionOutputModes.set(sessionId, mode as OutputMode);
          await sendText(context, `🔄 Output mode: **${mode}**`);
          return;
        }
      }
    }

    // Session cancel: verb = "cancel:<sessionId>"
    if (verb.startsWith("cancel:")) {
      const sessionId = verb.split(":")[1];
      if (sessionId) {
        const session = this.core.sessionManager.getSession(sessionId);
        if (session) {
          // Verify the action came from the session's conversation
          const conversationId = context.activity.conversation?.id;
          const sessionThread = session.threadIds.get("teams");
          if (sessionThread && conversationId && sessionThread !== conversationId) {
            log.warn({ sessionId, conversationId, sessionThread }, "[TeamsAdapter] cancel: conversation mismatch");
            return;
          }
          try {
            await session.destroy();
          } catch (err) {
            log.error({ err, sessionId }, "[TeamsAdapter] cancel: destroy failed");
            try {
              await sendText(context, "❌ Failed to cancel session");
            } catch { /* best effort */ }
            return;
          }
          try {
            await sendText(context, "❌ Session cancelled");
          } catch { /* best effort */ }
        }
      }
      return;
    }

    // Command button: verb = "cmd:<command>"
    if (verb.startsWith("cmd:")) {
      await setupCardActionCallbacks(context, this);
      return;
    }
  }

  // ─── Reaction handler ────────────────────────────────────────────────────

  /**
   * Handle message reactions (like, heart, laugh, surprised, sad, angry).
   *
   * Teams sends messageReaction activities when users react to bot messages.
   * We log them as engagement signals and emit an event so plugins can act
   * on them (e.g., aggregate feedback, adjust behavior).
   */
  private setupReactionHandler(): void {
    this.app.on("messageReaction" as any, async (context: any) => {
      try {
        const added = context.activity.reactionsAdded as Array<{ type: string }> | undefined;
        const removed = context.activity.reactionsRemoved as Array<{ type: string }> | undefined;
        const replyToId = context.activity.replyToId as string | undefined;
        const userId = context.activity.from?.id ?? "unknown";
        const userName = context.activity.from?.name;
        const conversationId = context.activity.conversation?.id;

        if (added?.length) {
          for (const reaction of added) {
            log.info(
              { reaction: reaction.type, replyToId, userId, userName, conversationId },
              "[TeamsAdapter] Reaction added",
            );

            // Map Teams reactions to sentiment signals
            const positive = ["like", "heart", "laugh"].includes(reaction.type);
            const negative = ["sad", "angry"].includes(reaction.type);

            // Emit reaction event so plugins (usage tracking, analytics) can consume it
            if (this.core.eventBus) {
              (this.core.eventBus as any).emit("teams:reaction", {
                type: reaction.type,
                sentiment: positive ? "positive" : negative ? "negative" : "neutral",
                replyToId,
                userId,
                userName,
                conversationId,
              });
            }
          }
        }

        if (removed?.length) {
          for (const reaction of removed) {
            log.debug(
              { reaction: reaction.type, replyToId, userId },
              "[TeamsAdapter] Reaction removed",
            );
          }
        }
      } catch (err) {
        log.warn({ err }, "[TeamsAdapter] Reaction handler error");
      }
    });
  }

  // ─── Task Module handlers (modal dialogs) ────────────────────────────────

  /**
   * Register task/fetch and task/submit invoke handlers for Teams modal dialogs.
   * Task Modules are triggered by Action.Submit with msteams.type = "task/fetch".
   */
  private setupTaskModuleHandlers(): void {
    // task/fetch — return the dialog card
    this.app.on("task.fetch" as any, async (context: any) => {
      try {
        const data = context.activity.value?.data as Record<string, unknown> | undefined;
        const dialogId = data?.dialogId as string | undefined;

        if (dialogId === "new-session") {
          const agents = this.core.agentManager.getAvailableAgents();
          const workspace = this.core.configManager.resolveWorkspace?.() ?? process.cwd();
          return buildNewSessionDialog(agents, workspace);
        }

        if (dialogId === "settings") {
          const sessionId = data?.sessionId as string | undefined;
          const session = sessionId ? this.core.sessionManager.getSession(sessionId) : undefined;
          const config = this.core.configManager.get();
          return buildSettingsDialog({
            defaultAgent: config.defaultAgent,
            workspace: this.core.configManager.resolveWorkspace?.(),
            outputMode: "medium",
            sessionId: sessionId ?? undefined,
            sessionAgent: session?.agentName,
            sessionModel: session?.getConfigByCategory?.("model")?.currentValue as string | undefined,
            sessionBypass: !!session?.clientOverrides?.bypassPermissions,
            sessionTts: session?.voiceMode,
          });
        }

        return buildDialogMessage("Unknown dialog");
      } catch (err) {
        log.error({ err }, "[TeamsAdapter] task.fetch error");
        return buildDialogMessage("Failed to load dialog");
      }
    });

    // task/submit — process form data from the dialog
    this.app.on("task.submit" as any, async (context: any) => {
      try {
        const data = context.activity.value?.data as Record<string, unknown> | undefined;
        const action = data?.dialogAction as string | undefined;

        if (action === "new-session") {
          const agentName = data?.agent as string;
          const workspace = data?.workspace as string;
          if (!agentName || !workspace) {
            return buildDialogMessage("Agent and workspace are required.");
          }

          // Validate agent name against registered agents
          const availableAgents = this.core.agentManager.getAvailableAgents();
          if (!availableAgents.some((a) => a.name === agentName)) {
            return buildDialogMessage(`Unknown agent: ${agentName}`);
          }

          // Validate workspace — must match the configured workspace to prevent path traversal
          const allowedWorkspace = this.core.configManager.resolveWorkspace?.() ?? process.cwd();
          if (workspace !== allowedWorkspace) {
            log.warn({ workspace, allowedWorkspace }, "[TeamsAdapter] task.submit: workspace mismatch");
            return buildDialogMessage("Invalid workspace path.");
          }

          try {
            const session = await this.core.sessionManager.createSession(
              "teams",
              agentName,
              allowedWorkspace,
              this.core.agentManager,
            );
            const threadId = await this.createSessionThread(session.id, session.name || agentName);
            session.threadId = threadId;
            session.threadIds.set("teams", threadId);

            // Store context for the new session
            const conversationId = context.activity.conversation?.id;
            if (conversationId) {
              this._sessionContexts.set(session.id, {
                context,
                isAssistant: false,
              });
            }

            return buildDialogMessage(`Session created with ${agentName} in ${workspace}`);
          } catch (err) {
            log.error({ err, agentName, workspace }, "[TeamsAdapter] task.submit new-session error");
            return buildDialogMessage(`Failed to create session: ${(err as Error).message}`);
          }
        }

        if (action === "save-settings") {
          const outputMode = data?.outputMode as string | undefined;
          const bypass = data?.bypass as string | undefined;
          const sessionId = data?.sessionId as string | undefined;

          // Verify the submitting user's conversation owns the target session
          const conversationId = context.activity.conversation?.id;
          const session = sessionId ? this.core.sessionManager.getSession(sessionId) : undefined;
          if (session && conversationId) {
            const sessionThread = session.threadIds.get("teams");
            if (sessionThread && sessionThread !== conversationId) {
              log.warn({ sessionId, conversationId, sessionThread }, "[TeamsAdapter] save-settings: conversation mismatch");
              return buildDialogMessage("You do not have access to this session.");
            }
          }

          if (outputMode && sessionId) {
            this.setSessionOutputMode(sessionId, outputMode as "low" | "medium" | "high");
          }
          if (bypass !== undefined && session) {
            session.clientOverrides = { ...session.clientOverrides, bypassPermissions: bypass === "true" };
          }

          return buildDialogMessage("Settings saved");
        }

        return buildDialogMessage("Unknown action");
      } catch (err) {
        log.error({ err }, "[TeamsAdapter] task.submit error");
        return buildDialogMessage("Failed to process");
      }
    });
  }

  /**
   * Handle a dialog request that arrived via Action.Submit (not invoke).
   * Since Action.Submit doesn't trigger the task/fetch invoke path,
   * we send the dialog card as a regular inline message instead.
   */
  private async handleInlineDialog(context: any, dialogId: string, data: Record<string, unknown>): Promise<void> {
    if (dialogId === "new-session") {
      const agents = this.core.agentManager.getAvailableAgents();
      const workspace = this.core.configManager.resolveWorkspace?.() ?? process.cwd();
      const dialog = buildNewSessionDialog(agents, workspace);
      const cardContent = (dialog.task as any)?.value?.card?.content;
      if (cardContent) {
        await sendCard(context, cardContent);
      }
      return;
    }

    if (dialogId === "settings") {
      const sessionId = data.sessionId as string | undefined;
      const session = sessionId ? this.core.sessionManager.getSession(sessionId) : undefined;
      const config = this.core.configManager.get();
      const dialog = buildSettingsDialog({
        defaultAgent: config.defaultAgent,
        workspace: this.core.configManager.resolveWorkspace?.(),
        outputMode: "medium",
        sessionId: sessionId ?? undefined,
        sessionAgent: session?.agentName,
        sessionModel: session?.getConfigByCategory?.("model")?.currentValue as string | undefined,
        sessionBypass: !!session?.clientOverrides?.bypassPermissions,
        sessionTts: session?.voiceMode,
      });
      const cardContent = (dialog.task as any)?.value?.card?.content;
      if (cardContent) {
        await sendCard(context, cardContent);
      }
      return;
    }

    await sendText(context, "Unknown dialog");
  }

  // ─── Card action handler (for Action.Execute / invoke activities) ─────

  private setupCardActionHandler(): void {
    this.app.on("card.action", (async (context: any): Promise<InvokeResponse<"adaptiveCard/action"> | undefined> => {
      await this.cardActionHandler(context);
      return { status: 200 };
    }) as any);
  }

  private async cardActionHandler(context: any): Promise<void> {
    try {
      const action = context.activity.value?.action;
      if (!action) return;

      const verb = action.verb as string;
      const data = action.data as Record<string, unknown> ?? {};

      if (!verb) return;

      await this.dispatchCardVerb(context, verb, data);
    } catch (err) {
      log.error({ err }, "[TeamsAdapter] card.action handler error");
    }
  }

  // ─── CommandRegistry dispatch ────────────────────────────────────────────

  private getCommandRegistry(): CommandRegistry | undefined {
    return this.core.lifecycleManager?.serviceRegistry?.get<CommandRegistry>("command-registry");
  }

  async handleCommand(
    text: string,
    context: TurnContext,
    sessionId: string | null,
    userId: string,
  ): Promise<void> {
    const registry = this.getCommandRegistry();
    if (!registry) return;

    const response = await registry.execute(text, {
      raw: "",
      sessionId,
      channelId: "teams",
      userId,
      reply: async (content: string) => {
        if (typeof content === "string") {
          try {
            await sendText(context, content);
          } catch (err) {
            log.warn({ err }, "[TeamsAdapter] handleCommand: reply failed");
          }
        }
      },
    });

    if (response.type !== "silent") {
      await this.renderCommandResponse(response, context);
    }
  }

  private async renderCommandResponse(response: CommandResponse, context: any): Promise<void> {
    const reply = async (text: string) => {
      await sendText(context, text);
    };

    switch (response.type) {
      case "text":
        await reply(response.text);
        break;
      case "error":
        await reply(`⚠️ ${response.message}`);
        break;
      case "menu": {
        const card = {
          type: "AdaptiveCard",
          version: "1.2",
          body: [
            { type: "TextBlock", text: response.title, weight: "Bolder", size: "Medium" },
            { type: "TextBlock", text: response.options.map((o: { label: string }) => `• ${o.label}`).join("\n"), wrap: true },
          ],
          actions: response.options.slice(0, 5).map((opt: { label: string; command: string }) => ({
            type: "Action.Submit",
            title: opt.label.slice(0, 20),
            data: { verb: `cmd:${opt.command}` },
          })),
        };
        await sendCard(context, card);
        break;
      }
      case "list": {
        const text = response.items.map((i: { label: string; detail?: string }) => `• **${i.label}**${i.detail ? ` — ${i.detail}` : ""}`).join("\n");
        await reply(`${response.title}\n${text}`);
        break;
      }
      case "confirm": {
        const card = {
          type: "AdaptiveCard",
          version: "1.2",
          body: [{ type: "TextBlock", text: response.question, wrap: true }],
          actions: [
            { type: "Action.Submit", title: "Yes", data: { verb: `cmd:${response.onYes}` } },
            { type: "Action.Submit", title: "No", data: { verb: "cmd:noop" } },
          ],
        };
        await sendCard(context, card);
        break;
      }
      case "silent":
        break;
      default:
        await reply(`⚠️ Unexpected response type: ${(response as Record<string, unknown>).type ?? "unknown"}`);
    }
  }

  // ─── Assistant ────────────────────────────────────────────────────────────

  private async setupAssistant(): Promise<void> {
    let threadId = this.teamsConfig.assistantThreadId ?? undefined;

    if (!threadId) {
      threadId = await this.createSessionThread("assistant", "Assistant") as string;
      this.teamsConfig.assistantThreadId = threadId;
      await this.core.configManager.save({
        channels: { teams: { assistantThreadId: threadId } },
      } as Parameters<typeof this.core.configManager.save>[0]);
      log.info({ threadId }, "[TeamsAdapter] Created assistant thread");
    }

    this.assistantInitializing = true;
    try {
      const { session, ready } = await spawnAssistant(this.core, threadId);
      this.assistantSession = session;

      // Guard ensures only one of timeout/ready acts on the buffer (prevents race)
      let settled = false;
      const settle = (replay: boolean) => {
        if (settled) return;
        settled = true;
        clearTimeout(timeout);
        this.assistantInitializing = false;
        const buffered = this._assistantInitBuffer.splice(0);
        if (replay) {
          for (const { sessionId: sid, content: msg } of buffered) {
            this.sendMessage(sid, msg).catch((err) => {
              log.warn({ err, sessionId: sid }, "[TeamsAdapter] Failed to replay buffered assistant message");
            });
          }
        }
      };

      const timeout = setTimeout(() => {
        log.warn("[TeamsAdapter] Assistant ready timeout — clearing initializing flag");
        settle(false); // discard stale messages on timeout
      }, 60_000);
      if (timeout.unref) timeout.unref();

      ready.then(
        () => settle(true),  // replay buffered messages on success
        (err) => {
          log.error({ err }, "[TeamsAdapter] Assistant ready promise rejected");
          settle(false); // discard buffer on failure
        },
      );
    } catch (err) {
      this.assistantInitializing = false;
      this._assistantInitBuffer.splice(0);
      log.error({ err }, "[TeamsAdapter] Failed to spawn assistant");
    }
  }

  async respawnAssistant(): Promise<void> {
    // Reset init state to prevent stale closures from a previous setupAssistant()
    // call (e.g., its 60s timeout) from interfering with the new spawn.
    this.assistantInitializing = false;
    this._assistantInitBuffer.splice(0);
    if (this.assistantSession) {
      try {
        await this.assistantSession.destroy();
      } catch { /* ignore */ }
      this.assistantSession = null;
    }
    await this.setupAssistant();
  }

  async restartAssistant(): Promise<void> {
    await this.respawnAssistant();
  }

  // ─── Typing indicator ────────────────────────────────────────────────────

  /** Send a typing indicator to the user. Non-critical — failures are silently ignored. */
  private sendTyping(context: TurnContext): void {
    sendActivity(context, { type: "typing" }).catch(() => {});
  }

  // ─── Bot token for proactive messaging ────────────────────────────────────

  /**
   * Acquire a bot framework token for proactive messaging via the MSA/AAD endpoint.
   * Required when posting to the Bot Connector REST API outside of a turn context.
   */
  private async acquireBotToken(): Promise<string | null> {
    if (this._botTokenCache && Date.now() < this._botTokenCache.expiresAt - 60_000) {
      return this._botTokenCache.token;
    }

    const appId = this.teamsConfig.botAppId;
    const appPassword = this.teamsConfig.botAppPassword;
    if (!appId || !appPassword) return null;

    try {
      // Use configured tenantId for single-tenant bots; fall back to botframework.com for multi-tenant
      const tenantForToken = this.teamsConfig.tenantId || "botframework.com";
      const response = await fetch(`https://login.microsoftonline.com/${tenantForToken}/oauth2/v2.0/token`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "client_credentials",
          client_id: appId,
          client_secret: appPassword,
          scope: "https://api.botframework.com/.default",
        }).toString(),
      });

      if (!response.ok) {
        log.warn({ status: response.status }, "[TeamsAdapter] Bot token acquisition failed");
        return null;
      }

      const data = (await response.json()) as { access_token: string; expires_in: number };
      this._botTokenCache = {
        token: data.access_token,
        expiresAt: Date.now() + data.expires_in * 1000,
      };
      return data.access_token;
    } catch (err) {
      log.warn({ err }, "[TeamsAdapter] Bot token acquisition error");
      return null;
    }
  }

  // ─── Retry helper — matches Telegram's retryWithBackoff and 429 handling ──

  /**
   * Validate that a serviceUrl is a trusted Bot Framework endpoint.
   * Prevents SSRF where a spoofed serviceUrl could redirect bot tokens.
   *
   * @see https://learn.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-api-reference
   */
  private static readonly TRUSTED_SERVICE_URL_PATTERNS = [
    /^https:\/\/[\w.-]+\.botframework\.com\b/i,
    /^https:\/\/[\w.-]+\.teams\.microsoft\.com\b/i,
    /^https:\/\/smba\.trafficmanager\.net\b/i,
    /^https:\/\/[\w.-]+\.botframework\.azure\.us\b/i,
    // Allow localhost for development
    /^https?:\/\/localhost(:\d+)?\b/,
    /^https?:\/\/127\.0\.0\.1(:\d+)?\b/,
  ];

  static isValidServiceUrl(url: string): boolean {
    return TeamsAdapter.TRUSTED_SERVICE_URL_PATTERNS.some((pattern) => pattern.test(url));
  }

  /** AI-generated content entity — attached to all outbound messages for the Teams "AI generated" badge */
  private static readonly AI_ENTITY = {
    type: "https://schema.org/Message",
    "@type": "Message",
    "@context": "https://schema.org",
    additionalType: ["AIGeneratedContent"],
  };

  /**
   * Send a Teams activity with exponential backoff retry on transient failures.
   * Handles HTTP 429 (rate limited), 502, 504 per Microsoft best practices.
   */
  private async sendActivityWithRetry(
    context: TurnContext,
    activity: Record<string, unknown>,
  ): Promise<unknown> {
    // Attach AI-generated content label to all message activities.
    // Clone the activity to avoid mutating the caller's object (and duplicating on retries).
    if (!activity.type || activity.type === "message") {
      const existing = (activity.entities as unknown[] | undefined) ?? [];
      // Skip if an AIGeneratedContent entity is already present (e.g., citation entities)
      const hasAiLabel = existing.some((e: any) =>
        Array.isArray(e?.additionalType) && e.additionalType.includes("AIGeneratedContent"),
      );
      if (!hasAiLabel) {
        activity = { ...activity, entities: [...existing, TeamsAdapter.AI_ENTITY] };
      }
    }

    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      try {
        return await sendActivity(context, activity);
      } catch (err: unknown) {
        const statusCode = (err as { statusCode?: number })?.statusCode;
        // Teams docs require retrying 412, 429, 502, and 504
        const isRetryable = statusCode === 412 || statusCode === 429 || statusCode === 502 || statusCode === 504;

        if (!isRetryable || attempt === MAX_RETRIES) throw err;

        // Parse Retry-After header if available, otherwise use exponential backoff + jitter
        const retryAfterRaw = (err as { headers?: Record<string, string> })?.headers?.["retry-after"];
        const retryAfterSec = retryAfterRaw ? parseInt(retryAfterRaw, 10) : NaN;
        const delayMs = !isNaN(retryAfterSec) && retryAfterSec > 0
          ? retryAfterSec * 1000
          : BASE_RETRY_DELAY * Math.pow(2, attempt) + Math.random() * 500;

        log.warn(
          { statusCode, attempt: attempt + 1, delayMs },
          "[TeamsAdapter] Rate limited or transient error, retrying",
        );
        await new Promise((r) => setTimeout(r, delayMs));
      }
    }
    throw new Error("unreachable");
  }

  // ─── Helper: resolve context ─────────────────────────────────────────────

  private resolveMode(sessionId: string): OutputMode {
    // Check session-level override
    const sessionMode = this._sessionOutputModes.get(sessionId);
    if (sessionMode) return sessionMode;

    // Session record
    const record = this.core.sessionManager.getSessionRecord(sessionId);
    if (record?.outputMode) {
      const m = record.outputMode;
      if (m === "low" || m === "medium" || m === "high") return m as OutputMode;
    }

    // Adapter config — check teams-specific setting
    const config = this.core.configManager.get() as Record<string, unknown>;
    const channels = config.channels as Record<string, Record<string, unknown>> | undefined;
    const adapterMode = channels?.teams?.outputMode;
    if (adapterMode === "low" || adapterMode === "medium" || adapterMode === "high") {
      return adapterMode as OutputMode;
    }

    // Global
    if (config.outputMode === "low" || config.outputMode === "medium" || config.outputMode === "high") {
      return config.outputMode as OutputMode;
    }

    return "medium";
  }

  // ─── sendMessage ─────────────────────────────────────────────────────────

  /**
   * Primary outbound dispatch — routes agent messages to Teams.
   *
   * Wraps the base class `sendMessage` in a per-session promise chain (_dispatchQueues)
   * so concurrent events fired from SessionBridge are serialized and delivered in order,
   * preventing fast handlers from overtaking slower ones (matches Telegram pattern).
   *
   * Context is NOT deleted after dispatch — it persists from the inbound message handler
   * and is available for the entire session lifetime, avoiding the race condition where
   * async handlers lose their context mid-execution.
   */
  async sendMessage(sessionId: string, content: OutgoingMessage): Promise<void> {
    // Buffer messages during assistant initialization instead of dropping them
    if (
      this.assistantInitializing &&
      this.assistantSession &&
      sessionId === this.assistantSession.id
    ) {
      if (this._assistantInitBuffer.length < TeamsAdapter.MAX_INIT_BUFFER) {
        this._assistantInitBuffer.push({ sessionId, content });
      } else {
        log.warn({ sessionId }, "[TeamsAdapter] Assistant init buffer full, dropping message");
      }
      return;
    }

    // Look up context by sessionId first, then by threadId (for newly-created sessions
    // where context was stored under threadId before the session existed).
    // Multiple fallback paths ensure the first response to a new session is not dropped.
    let ctx = this._sessionContexts.get(sessionId);
    if (!ctx) {
      // Try in-memory session's threadId
      const session = this.core.sessionManager.getSession(sessionId);
      const threadId = session?.threadId;
      if (threadId) {
        ctx = this._sessionContexts.get(threadId);
      }
      // Try stored session record's threadId (covers async session creation)
      if (!ctx) {
        const record = this.core.sessionManager.getSessionRecord(sessionId);
        const recordThreadId = (record?.platform as Record<string, unknown>)?.threadId as string | undefined;
        if (recordThreadId) {
          ctx = this._sessionContexts.get(recordThreadId);
        }
      }
      // Promote to sessionId key for future lookups
      if (ctx) this._sessionContexts.set(sessionId, ctx);
    }
    if (!ctx) {
      // No live TurnContext — for terminal events (error, session_end), attempt
      // proactive delivery so the user isn't left without feedback
      if (content.type === "error" || content.type === "session_end") {
        log.warn({ sessionId, type: content.type }, "[TeamsAdapter] sendMessage: no context, attempting proactive delivery");
        const text = content.type === "error"
          ? `❌ **Error:** ${content.text}`
          : "✅ **Done**";
        await this.sendNotification({
          sessionId,
          type: content.type === "error" ? "error" : "completed",
          summary: text,
        });
      } else {
        log.warn({ sessionId, type: content.type }, "[TeamsAdapter] sendMessage: no context for session, skipping");
      }
      return;
    }

    // Serialize dispatch per session to preserve event ordering.
    // Read + write the queue entry atomically (synchronous) so concurrent callers
    // always chain on the latest promise, preventing parallel execution.
    const prev = this._dispatchQueues.get(sessionId) ?? Promise.resolve();
    const next = prev.then(async () => {
      try {
        await super.sendMessage(sessionId, content);
      } catch (err) {
        log.warn({ err, sessionId, type: content.type }, "[TeamsAdapter] Dispatch error");
      }
    });
    // Set immediately — before any await — so the next concurrent caller sees this entry
    this._dispatchQueues.set(sessionId, next);
    await next;
    // Replace settled chain with a fresh resolved promise to prevent unbounded
    // closure growth for long-lived sessions. Only replace if no new work was
    // chained while we were awaiting (i.e., the entry still points to `next`).
    if (this._dispatchQueues.get(sessionId) === next) {
      this._dispatchQueues.set(sessionId, Promise.resolve());
    }
  }

  // ─── Handler overrides ───────────────────────────────────────────────────

  protected async handleThought(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    // Thoughts are not sent as messages in Teams — buffered and displayed via plan
    void sessionId;
    void content;
  }

  protected async handleText(sessionId: string, content: OutgoingMessage): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const { context } = ctx;
    // Send typing indicator on first text chunk (before draft exists)
    if (!this.draftManager.hasDraft(sessionId)) {
      this.sendTyping(context);
    }
    const draft = this.draftManager.getOrCreate(sessionId, context);
    if (content.text) draft.append(content.text);
  }

  protected async handleToolCall(sessionId: string, content: OutgoingMessage, verbosity: DisplayVerbosity): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const { context, isAssistant } = ctx;
    this.sendTyping(context);
    await this.draftManager.finalize(sessionId, context, isAssistant);
    try {
      const meta = (content.metadata ?? {}) as Partial<ToolCallMeta>;
      const cardData = renderToolCallCard({
        id: meta.id ?? "",
        name: meta.name ?? content.text ?? "Tool",
        kind: meta.kind,
        status: meta.status,
        rawInput: meta.rawInput,
        content: meta.content,
        displaySummary: meta.displaySummary as string | undefined,
        displayTitle: meta.displayTitle as string | undefined,
        displayKind: meta.displayKind as string | undefined,
        viewerLinks: meta.viewerLinks as { file?: string; diff?: string } | undefined,
        viewerFilePath: meta.viewerFilePath as string | undefined,
      }, verbosity);
      const card = { type: "AdaptiveCard", version: "1.2", ...cardData };

      // Build citation entities for file references (hover popup with source info)
      const citationSources: Array<{ name: string; url: string; abstract?: string }> = [];
      const filePath = meta.viewerFilePath as string | undefined;
      const links = meta.viewerLinks as { file?: string; diff?: string } | undefined;
      if (filePath && links?.file) {
        citationSources.push({ name: filePath.split("/").pop() || filePath, url: links.file, abstract: `Source: ${filePath}` });
      }
      if (filePath && links?.diff) {
        citationSources.push({ name: `${filePath.split("/").pop() || filePath} (diff)`, url: links.diff, abstract: `Changes to ${filePath}` });
      }
      const citationEntities = buildCitationEntities(citationSources);

      // Citations require [N] markers in the text field — Teams ignores citation entities
      // on card-only activities with no text anchor. Add a text field with markers.
      const citationText = citationSources.length > 0
        ? citationSources.map((s, i) => `[${i + 1}]`).join(" ")
        : undefined;

      await this.sendActivityWithRetry(context, {
        ...(citationText ? { text: citationText } : {}),
        attachments: [CardFactory.adaptiveCard(card as Record<string, unknown>)],
        ...(citationEntities.length > 0 ? { entities: citationEntities } : {}),
      });
    } catch (err) {
      log.error({ err, sessionId }, "[TeamsAdapter] handleToolCall: sendActivity failed");
    }
  }

  protected async handleToolUpdate(sessionId: string, content: OutgoingMessage, verbosity: DisplayVerbosity): Promise<void> {
    // Only render tool updates in high verbosity mode (matches Telegram's tracker behavior)
    if (verbosity !== "high") return;
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const { context } = ctx;
    try {
      const rendered = this.renderer.renderToolUpdate(content, verbosity);
      await this.sendActivityWithRetry(context, { text: rendered.body });
    } catch (err) {
      log.warn({ err, sessionId }, "[TeamsAdapter] handleToolUpdate: sendActivity failed");
    }
  }

  protected async handlePlan(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const { context } = ctx;
    const entries = (content.metadata as { entries?: PlanEntry[] })?.entries ?? [];
    const mode = this.resolveMode(sessionId);
    const cardData = renderPlanCard(entries, mode);
    const card = { type: "AdaptiveCard", version: "1.2", ...cardData };
    try {
      await this.sendActivityWithRetry(context, { attachments: [CardFactory.adaptiveCard(card as Record<string, unknown>)] });
    } catch (err) {
      log.error({ err, sessionId }, "[TeamsAdapter] handlePlan: sendActivity failed");
    }
  }

  protected async handleUsage(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const { context, isAssistant } = ctx;
    await this.draftManager.finalize(sessionId, context, isAssistant);
    const meta = content.metadata as { tokensUsed?: number; contextSize?: number; cost?: number; duration?: number } | undefined;
    const mode = this.resolveMode(sessionId);
    const { body } = renderUsageCard(meta ?? {}, mode);
    const card = { type: "AdaptiveCard" as const, version: "1.2" as const, body };
    try {
      // Feedback buttons on usage/completion messages — Teams shows thumbs up/down
      await this.sendActivityWithRetry(context, {
        attachments: [CardFactory.adaptiveCard(card as Record<string, unknown>)],
        channelData: { feedbackLoop: { type: "default" } },
      });
    } catch (err) {
      log.error({ err, sessionId }, "[TeamsAdapter] handleUsage: sendActivity failed");
    }

    // Notify completion in notification channel (matches Telegram's notification pattern)
    if (this.notificationChannelId && sessionId !== this.assistantSession?.id) {
      const sess = this.core.sessionManager.getSession(sessionId);
      const name = sess?.name || "Session";
      void this.sendNotification({
        sessionId,
        sessionName: name,
        type: "completed",
        summary: "Task completed",
      });
    }
  }

  /** Suggested quick-reply actions (Teams restricts these to 1:1 personal chat only) */
  private static readonly QUICK_ACTIONS = {
    suggestedActions: {
      actions: [
        { type: "imBack", title: "➕ New Session", value: "/new" },
        { type: "imBack", title: "📊 Status", value: "/status" },
        { type: "imBack", title: "📋 Sessions", value: "/sessions" },
        { type: "imBack", title: "📋 Menu", value: "/menu" },
      ],
    },
  };

  /** Return QUICK_ACTIONS only if the conversation is 1:1 personal chat (Teams requirement) */
  private getQuickActions(context: TurnContext): Record<string, unknown> {
    const convType = (context.activity as Record<string, unknown>).conversation as Record<string, unknown> | undefined;
    if (convType?.conversationType === "personal") {
      return TeamsAdapter.QUICK_ACTIONS;
    }
    return {};
  }

  /**
   * Clean up all per-session state (contexts, drafts, dispatch queues, output modes).
   * Removes both sessionId and threadId entries from _sessionContexts to prevent leaks.
   */
  private cleanupSessionState(sessionId: string): void {
    // Find and remove the threadId entry that may also reference this session's context.
    // First try the stored threadId on the context entry itself (reliable even if the
    // session has already been removed from the session manager).
    const entry = this._sessionContexts.get(sessionId);
    const storedThreadId = entry?.threadId;
    if (storedThreadId && storedThreadId !== sessionId) {
      this._sessionContexts.delete(storedThreadId);
    }
    // Fallback: also check session manager and session record in case the context entry
    // was already removed or the threadId wasn't stored.
    const session = this.core.sessionManager.getSession(sessionId);
    const threadId = session?.threadId;
    if (threadId && threadId !== sessionId && threadId !== storedThreadId) {
      this._sessionContexts.delete(threadId);
    }
    const record = this.core.sessionManager.getSessionRecord(sessionId);
    const recordThreadId = (record?.platform as Record<string, unknown>)?.threadId as string | undefined;
    if (recordThreadId && recordThreadId !== sessionId && recordThreadId !== threadId && recordThreadId !== storedThreadId) {
      this._sessionContexts.delete(recordThreadId);
    }

    this._sessionContexts.delete(sessionId);
    this._sessionOutputModes.delete(sessionId);
    this._dispatchQueues.delete(sessionId);
    this.draftManager.cleanup(sessionId);
  }

  protected async handleSessionEnd(sessionId: string, _content: OutgoingMessage): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const { context, isAssistant } = ctx;
    await this.draftManager.finalize(sessionId, context, isAssistant);
    this.cleanupSessionState(sessionId);
    try {
      await this.sendActivityWithRetry(context, {
        text: "✅ **Done**",
        channelData: { feedbackLoop: { type: "default" } },
        ...this.getQuickActions(context),
      });
    } catch { /* best effort */ }
  }

  protected async handleError(sessionId: string, content: OutgoingMessage): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const { context, isAssistant } = ctx;
    await this.draftManager.finalize(sessionId, context, isAssistant);
    this.cleanupSessionState(sessionId);
    try {
      await this.sendActivityWithRetry(context, {
        text: `❌ **Error:** ${content.text}`,
        ...this.getQuickActions(context),
      });
    } catch { /* best effort */ }
  }

  protected async handleAttachment(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (!content.attachment) return;
    const { attachment } = content;
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const { context, isAssistant } = ctx;

    // Strip TTS markers from the draft BEFORE finalizing — finalize() deletes
    // the draft from the map, so getDraft() would return undefined after it.
    if (attachment.type === "audio") {
      const draft = this.draftManager.getDraft(sessionId);
      if (draft) {
        await draft.stripPattern(/\[TTS\][\s\S]*?\[\/TTS\]/g).catch((err) => {
          log.warn({ err, sessionId }, "[TeamsAdapter] handleAttachment: stripPattern failed");
        });
      }
    }

    await this.draftManager.finalize(sessionId, context, isAssistant);

    if (isAttachmentTooLarge(attachment.size)) {
      log.warn({ sessionId, fileName: attachment.fileName, size: attachment.size }, "[TeamsAdapter] File too large");
      try {
        await this.sendActivityWithRetry(context, {
          text: `⚠️ File too large to send (${Math.round(attachment.size / 1024 / 1024)}MB): ${attachment.fileName}`,
        });
      } catch { /* best effort */ }
      return;
    }

    try {
      // Upload to OneDrive via Graph API if available, get a sharing URL
      const shareUrl = await uploadFileViaGraph(
        this.graphClient,
        sessionId,
        attachment.filePath,
        attachment.fileName,
        attachment.mimeType,
      );

      const card = buildFileAttachmentCard(attachment.fileName, attachment.size, attachment.mimeType, shareUrl ?? undefined);
      await this.sendActivityWithRetry(context, { attachments: [CardFactory.adaptiveCard(card as Record<string, unknown>)] });
    } catch (err) {
      log.error({ err, sessionId, fileName: attachment.fileName }, "[TeamsAdapter] Failed to send attachment");
    }
  }

  protected async handleSystem(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (!content.text) return;
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const { context } = ctx;
    try {
      await this.sendActivityWithRetry(context, { text: content.text });
    } catch { /* best effort */ }
  }

  protected async handleModeChange(sessionId: string, content: OutgoingMessage): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const renderer = this.renderer as TeamsRenderer;
    const rendered = renderer.renderModeChange(content);
    try {
      await this.sendActivityWithRetry(ctx.context, { text: rendered.body });
    } catch { /* best effort */ }
  }

  protected async handleConfigUpdate(sessionId: string, content: OutgoingMessage): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const renderer = this.renderer as TeamsRenderer;
    const rendered = renderer.renderConfigUpdate(content);
    try {
      await this.sendActivityWithRetry(ctx.context, { text: rendered.body });
    } catch { /* best effort */ }
  }

  protected async handleModelUpdate(sessionId: string, content: OutgoingMessage): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const renderer = this.renderer as TeamsRenderer;
    const rendered = renderer.renderModelUpdate(content);
    try {
      await this.sendActivityWithRetry(ctx.context, { text: rendered.body });
    } catch { /* best effort */ }
  }

  protected async handleUserReplay(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (!content.text) return;
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    try {
      await this.sendActivityWithRetry(ctx.context, { text: content.text });
    } catch { /* best effort */ }
  }

  protected async handleResource(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (!content.text) return;
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    try {
      await this.sendActivityWithRetry(ctx.context, { text: content.text });
    } catch { /* best effort */ }
  }

  protected async handleResourceLink(sessionId: string, content: OutgoingMessage): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const rawUrl = (content.metadata as Record<string, unknown>)?.url as string | undefined;
    const rawName = (content.metadata as Record<string, unknown>)?.name as string | undefined;
    // Only allow http/https URLs to prevent javascript: or data: scheme injection
    const url = rawUrl && /^https?:\/\//i.test(rawUrl) ? rawUrl : undefined;
    // Sanitize name to prevent markdown injection — strip characters that break link syntax
    const name = rawName?.replace(/[\[\]\(\)]/g, "") || undefined;
    const text = url ? `📎 [${name || url}](${url})` : content.text;
    try {
      await this.sendActivityWithRetry(ctx.context, { text });
    } catch { /* best effort */ }
  }

  // ─── sendPermissionRequest ──────────────────────────────────────────────

  async sendPermissionRequest(sessionId: string, request: PermissionRequest): Promise<void> {
    const session = this.core.sessionManager.getSession(sessionId);
    if (!session) {
      log.warn({ sessionId }, "[TeamsAdapter] sendPermissionRequest: session not found");
      return;
    }
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) {
      log.warn({ sessionId }, "[TeamsAdapter] sendPermissionRequest: no context");
      return;
    }
    await this.permissionHandler.sendPermissionRequest(session, request, ctx.context);
  }

  // ─── sendNotification ──────────────────────────────────────────────────

  async sendNotification(notification: NotificationMessage): Promise<void> {
    const typeIcon: Record<string, string> = {
      completed: "✅", error: "❌", permission: "🔐", input_required: "💬", budget_warning: "⚠️",
    };

    const icon = typeIcon[notification.type] ?? "ℹ️";
    const name = notification.sessionName ? ` **${notification.sessionName}**` : "";
    let text = `${icon}${name}: ${notification.summary}`;
    if (notification.deepLink) {
      text += `\n${notification.deepLink}`;
    }

    // Proactive messaging via stored conversation reference + bot token.
    // We do NOT use a stored TurnContext here — TurnContexts are scoped to a
    // single HTTP request/response cycle and go stale after the turn ends.
    if (this.notificationChannelId) {
      const ref = this.conversationStore.get(this.notificationChannelId) ?? this.conversationStore.getAny();
      if (ref && TeamsAdapter.isValidServiceUrl(ref.serviceUrl)) {
        try {
          const botToken = await this.acquireBotToken();
          if (botToken) {
            const controller = new AbortController();
            const timeout = setTimeout(() => controller.abort(), 10_000);
            try {
              const response = await fetch(`${ref.serviceUrl}/v3/conversations/${ref.conversationId}/activities`, {
                method: "POST",
                headers: {
                  "Content-Type": "application/json",
                  "Authorization": `Bearer ${botToken}`,
                },
                body: JSON.stringify({
                  type: "message",
                  text,
                  from: { id: ref.botId, name: ref.botName },
                }),
                signal: controller.signal,
              });
              if (response.ok) return;
              log.warn({ status: response.status }, "[TeamsAdapter] Proactive notification failed");
            } finally {
              clearTimeout(timeout);
            }
          }
        } catch (err) {
          log.warn({ err }, "[TeamsAdapter] Proactive notification error");
        }
      }
    }

    // Session-specific context fallback — last resort, may fail if the
    // TurnContext's HTTP response stream has closed since the turn ended.
    if (notification.sessionId) {
      const ctx = this._sessionContexts.get(notification.sessionId);
      if (ctx) {
        try {
          await sendText(ctx.context, text);
          return;
        } catch (err) {
          log.debug({ err, sessionId: notification.sessionId }, "[TeamsAdapter] Session context fallback failed (context may be stale)");
        }
      }
    }

    log.debug({ type: notification.type, sessionName: notification.sessionName }, "[TeamsAdapter] sendNotification: no delivery path available");
  }

  // ─── createSessionThread ─────────────────────────────────────────────────

  /**
   * Create a new conversation thread for a session.
   *
   * Attempts to create a real Teams channel conversation via the Bot Framework
   * connector API. If that fails (e.g., missing permissions, no stored conversation
   * reference), falls back to using the existing conversation ID as thread context.
   */
  async createSessionThread(sessionId: string, name: string): Promise<string> {
    // Try to create a real Teams conversation thread
    const ref = this.conversationStore.getAny();
    if (ref && TeamsAdapter.isValidServiceUrl(ref.serviceUrl)) {
      try {
        const botToken = await this.acquireBotToken();
        if (!botToken) throw new Error("No bot token available");
        const createUrl = `${ref.serviceUrl}/v3/conversations`;
        const response = await fetch(createUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json", "Authorization": `Bearer ${botToken}` },
          body: JSON.stringify({
            isGroup: false,
            bot: { id: ref.botId, name: ref.botName },
            tenantId: ref.tenantId,
            activity: {
              type: "message",
              text: `**${name.replace(/[*_~`[\]()\\]/g, "")}** — New session started`,
            },
            channelData: {
              channel: { id: this.teamsConfig.channelId },
              tenant: { id: ref.tenantId },
            },
          }),
        });

        if (response.ok) {
          const data = (await response.json()) as { id: string };
          const threadId = data.id;

          const session = this.core.sessionManager.getSession(sessionId);
          if (session) session.threadId = threadId;

          const record = this.core.sessionManager.getSessionRecord(sessionId);
          if (record) {
            await this.core.sessionManager.patchRecord(sessionId, {
              platform: { ...record.platform, threadId },
            });
          }

          log.info({ sessionId, threadId, name }, "[TeamsAdapter] Created real Teams conversation thread");
          return threadId;
        }

        log.warn({ status: response.status }, "[TeamsAdapter] createConversation failed, using fallback");
      } catch (err) {
        log.warn({ err }, "[TeamsAdapter] createConversation error, using fallback");
      }
    }

    // Fallback: use the configured channel as the conversation context
    const threadId = this.teamsConfig.channelId || `teams-${sessionId}-${Date.now()}`;
    const session = this.core.sessionManager.getSession(sessionId);
    if (session) session.threadId = threadId;

    const record = this.core.sessionManager.getSessionRecord(sessionId);
    if (record) {
      await this.core.sessionManager.patchRecord(sessionId, {
        platform: { ...record.platform, threadId },
      });
    }
    return threadId;
  }

  /**
   * Rename a session thread. This is a no-op for Teams — the Teams API does not
   * support renaming channel conversations. Renaming a group chat requires
   * Graph API with Chat.ReadWrite.All permission, which most bot registrations
   * don't have. The new name is stored in the session record for display purposes.
   */
  async renameSessionThread(sessionId: string, newName: string): Promise<void> {
    const record = this.core.sessionManager.getSessionRecord(sessionId);
    if (!record) return;

    // Persist the name in the session record even though Teams can't be updated
    try {
      await this.core.sessionManager.patchRecord(sessionId, {
        platform: { ...record.platform, displayName: newName },
      });
    } catch { /* best effort */ }

    log.debug({ sessionId, newName }, "[TeamsAdapter] renameSessionThread — name stored locally (Teams API does not support conversation rename)");
  }

  async deleteSessionThread(sessionId: string): Promise<void> {
    const record = this.core.sessionManager.getSessionRecord(sessionId);
    if (!record) return;
    const threadId = record.platform?.threadId as string | undefined;
    if (!threadId) return;

    // Clean up local state — actual Teams conversation deletion requires Graph API
    try {
      await this.core.sessionManager.patchRecord(sessionId, {
        platform: { ...record.platform, threadId: undefined },
      });
    } catch (err) {
      log.warn({ err, sessionId }, "[TeamsAdapter] deleteSessionThread: failed to clear threadId");
    }
    log.info({ sessionId, threadId }, "[TeamsAdapter] deleteSessionThread — cleanup done, Graph API delete not called");
  }

  // ─── Public helpers ─────────────────────────────────────────────────────

  getChannelId(): string {
    return this.teamsConfig.channelId;
  }

  getTeamId(): string {
    return this.teamsConfig.teamId;
  }

  getAssistantSessionId(): string | null {
    return this.assistantSession?.id ?? null;
  }

  getAssistantThreadId(): string | null {
    return this.teamsConfig.assistantThreadId ?? null;
  }

  setSessionOutputMode(sessionId: string, mode: "low" | "medium" | "high"): void {
    this._sessionOutputModes.set(sessionId, mode as OutputMode);
  }
}