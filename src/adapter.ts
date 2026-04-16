import { App } from "@microsoft/teams.apps";
import { BotBuilderPlugin } from "@microsoft/teams.botbuilder";
import { MemoryStorage } from "@microsoft/agents-hosting";
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
  MessagingAdapterConfig,
  FileServiceInterface,
  CommandResponse,
} from "@openacp/plugin-sdk";
import { log, MessagingAdapter, BaseRenderer } from "@openacp/plugin-sdk";
import type { CommandRegistry, ToolCallMeta, IRenderer } from "@openacp/plugin-sdk";
import type { TeamsChannelConfig } from "./types.js";
import { DEFAULT_BOT_PORT } from "./types.js";
import { SessionMessageManager, buildLevel1, buildLevel2, escapeMd } from "./message-composer.js";
import { ConversationRateLimiter } from "./rate-limiter.js";
import { PermissionHandler } from "./permissions.js";
import { handleCommand, setupCardActionCallbacks, SLASH_COMMANDS } from "./commands/index.js";
import { spawnAssistant } from "./assistant.js";
import { downloadTeamsFile, isAttachmentTooLarge, buildFileAttachmentCard, uploadFileViaGraph } from "./media.js";
import { GraphFileClient } from "./graph.js";
import { ConversationStore } from "./conversation-store.js";
import { sendText, sendCard } from "./send-utils.js";
import { formatTokens, formatToolSummary } from "./formatting.js";
import type { OutputMode } from "./activity.js";

export class TeamsAdapter extends MessagingAdapter {
  readonly name = "teams";
  readonly renderer: IRenderer = new BaseRenderer();
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
  private rateLimiter: ConversationRateLimiter;
  private composer: SessionMessageManager;
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

  /** Track processed activity IDs to handle Teams 15-second retry deduplication */
  private _processedActivities = new Map<string, number>();
  private _processedCleanupTimer?: ReturnType<typeof setInterval>;

  /** Per-session active tool entry ID for updateToolResult dispatch */
  private _toolEntryIds = new Map<string, string>();

  /** Messages buffered during assistant initialization — replayed once ready. Capped to prevent unbounded growth. */
  private static readonly MAX_INIT_BUFFER = 50;
  private _assistantInitBuffer: Array<{ sessionId: string; content: OutgoingMessage }> = [];

  /** Bot token cache for proactive messaging via connector REST API */
  private _botTokenCache?: { token: string; expiresAt: number };

  constructor(core: OpenACPCore, config: TeamsChannelConfig) {
    super(
      { configManager: core.configManager },
      // Teams measures message size in bytes (100KB limit, 80KB safe threshold).
      // Teams supports ~28k chars per message. Use 25000 to leave room for markdown overhead.
      // where each char can be 3-4 bytes, plus activity envelope overhead.
      { ...config as unknown as Record<string, unknown>, maxMessageLength: 25000, enabled: config.enabled ?? true } as MessagingAdapterConfig,
    );
    this.core = core;
    this.teamsConfig = config;
    this.rateLimiter = new ConversationRateLimiter();
    this.composer = new SessionMessageManager(this.rateLimiter, () => this.acquireBotToken());
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

  private _started = false;

  async start(): Promise<void> {
    if (this._started) {
      log.warn("[TeamsAdapter] Already started, skipping duplicate start()");
      return;
    }
    this._started = true;
    log.info("[TeamsAdapter] Starting...");

    try {
      this.notificationChannelId = this.teamsConfig.notificationChannelId ?? undefined;

      this.permissionHandler = new PermissionHandler(
        (sessionId) => this.core.sessionManager.getSession(sessionId),
        (notification) => this.sendNotification(notification),
        this.composer,
      );

      this.setupMessageHandler();
      this.setupCardActionHandler();

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
    this._processedActivities.clear();
    this.rateLimiter.destroy();
    this.conversationStore.destroy();
    this.permissionHandler.dispose();

    await this.app.stop();
    this._started = false;
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
      log.debug({ rawText: rawActivityText.slice(0, 100), cleanText: text.slice(0, 100), activityType: context.activity.type }, "[TeamsAdapter] Incoming activity");

      const userId = context.activity.from?.id ?? "unknown";
      // Use conversation.id as the thread discriminator — NOT activity.channelId
      // which is always "msteams" for Teams. Conversation ID uniquely identifies
      // the 1:1, group chat, or channel thread the message came from.
      // conversation.id may include a messageid suffix when received via group chat
      // or channel (e.g. "19:xxx@thread.tacv2;messageid=1776214235066"). Strip it before
      // checking the allowlist so the bare thread ID matches what's stored in config.
      const conversationId = String(context.activity.conversation?.id ?? "unknown").split(";")[0];
      const threadId = conversationId;

      // Security: only respond in configured channels
      const allowedChannelIds = this.teamsConfig.allowedChannelIds ?? [this.teamsConfig.channelId];
      if (!allowedChannelIds.includes(conversationId)) {
        log.debug({ conversationId, allowedChannelIds }, "[TeamsAdapter] Channel not in allowlist, ignoring");
        return;
      }

      // Security: enforce tenant isolation for single-tenant bots
      const isSingleTenant = this.teamsConfig.tenantId && this.teamsConfig.tenantId !== "botframework.com";
      if (isSingleTenant) {
        const incomingTenant = context.activity.conversation?.tenantId;
        if (!incomingTenant) {
          log.warn({ conversationId }, "[TeamsAdapter] Missing tenantId in activity — rejected (single-tenant mode)");
          return;
        }
        if (incomingTenant !== this.teamsConfig.tenantId) {
          log.warn({ incomingTenant, configuredTenant: this.teamsConfig.tenantId }, "[TeamsAdapter] Rejected message from unexpected tenant");
          return;
        }
      }

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

        // Don't cleanup the existing card here — if the agent is still working on
        // the previous prompt, its remaining events need the current card. The card
        // finalizes naturally via handleSessionEnd or handleError. The new prompt is
        // queued by core and its events will create a fresh card after the current
        // turn completes.

        const existingSessionBeforeSend = this.core.sessionManager.getSessionByThread("teams", threadId);
        if (!existingSessionBeforeSend) {
          const defaultAgent = (this.core.configManager.get() as Record<string, unknown>)?.defaultAgent as string ?? "claude";
          log.info({ threadId, text: messageText.slice(0, 50), defaultAgent }, "[TeamsAdapter] No session — auto-creating via /new");
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
                // Fire and forget — response comes back via sendMessage/SessionBridge
                this.core.handleMessage({
                  channelId: "teams",
                  threadId,
                  userId,
                  text: messageText,
                  ...(attachments.length > 0 ? { attachments } : {}),
                }).catch(err => log.error({ err }, "[TeamsAdapter] handleMessage failed"));
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

        // Fire and forget — response comes back via sendMessage/SessionBridge.
        // Don't await: a hung prompt (token exhaustion, etc.) would block all
        // subsequent messages in this conversation.
        this.core.handleMessage({
          channelId: "teams",
          threadId,
          userId,
          text: messageText,
          ...(attachments.length > 0 ? { attachments } : {}),
        }).catch(err => log.error({ err }, "[TeamsAdapter] handleMessage failed"));
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
      // Inline dialog form submissions (dialogAction from inline wizard cards)
      const dialogAction = data.dialogAction as string | undefined;
      if (dialogAction) {
        await this.handleDialogAction(context, dialogAction, data);
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
          const conversationId = (context.activity.conversation?.id as string | undefined)?.split(";")[0];
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
          const conversationId = (context.activity.conversation?.id as string | undefined)?.split(";")[0];
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

    // Inline dialog form submission: verb = "dialog:<action>"
    if (verb.startsWith("dialog:")) {
      const action = verb.slice(7);
      await this.handleDialogAction(context, action, data);
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
        const conversationId = (context.activity.conversation?.id as string | undefined)?.split(";")[0];

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


  /**
   * Create a session in the background — sends status updates to the conversation.
   * Runs async without blocking the invoke response.
   */
  private createSessionInBackground(context: any, agentName: string, workspace: string): void {
    const conversationId = (context.activity.conversation?.id as string | undefined)?.split(";")[0];
    (async () => {
      try {
        // Use core.createSession (not sessionManager.createSession) so the
        // SessionBridge is connected — without it, agent responses never
        // reach the adapter's sendMessage.
        const session = await (this.core as any).createSession({
          channelId: "teams",
          agentName,
          workingDirectory: workspace,
          threadId: conversationId,
          createThread: !conversationId,
        });

        // Ensure the thread is linked for DM conversations
        if (conversationId) {
          session.threadIds.set("teams", conversationId);
        }

        // Store context so the adapter can send responses to this conversation
        this._sessionContexts.set(session.id, { context, isAssistant: false, threadId: conversationId });

        const successCard = TeamsAdapter.buildNotificationCard("✅", "Session created", `${agentName} · ${workspace}`);
        await sendCard(context, successCard);
      } catch (err) {
        log.error({ err, agentName, workspace }, "[TeamsAdapter] createSessionInBackground error");
        try {
          const errorCard = TeamsAdapter.buildNotificationCard("❌", "Failed", (err as Error).message);
          await sendCard(context, errorCard);
        } catch { /* best effort */ }
      }
    })();
  }

  /**
   * Handle form submissions from inline wizard cards (dialogAction payloads).
   * These come from Action.Execute buttons on cards rendered directly in chat.
   */
  private async handleDialogAction(context: any, action: string, data: Record<string, unknown>): Promise<void> {
    if (action === "new-session") {
      const agentName = data.agent as string;
      const workspace = data.workspace as string;
      if (!agentName || !workspace) {
        await sendText(context, "❌ Agent and workspace are required.");
        return;
      }

      const availableAgents = this.core.agentManager.getAvailableAgents();
      if (!availableAgents.some((a) => a.name === agentName)) {
        await sendText(context, `❌ Unknown agent: ${agentName}`);
        return;
      }

      // Destroy any existing session in this conversation before creating a new one
      const conversationId = (context.activity.conversation?.id as string | undefined)?.split(";")[0];
      if (conversationId) {
        const existing = this.core.sessionManager.getSessionByThread("teams", conversationId);
        if (existing) {
          await this.composer.finalize(existing.id);
          try { await existing.destroy(); } catch { /* best effort */ }
        }
      }

      // Send acknowledgment immediately, then create session in the background.
      // Session creation spawns an agent process (~30s) which would timeout the invoke.
      const creatingCard = TeamsAdapter.buildNotificationCard("🔧", "Creating session", agentName);
      await sendCard(context, creatingCard);
      this.createSessionInBackground(context, agentName, workspace);
      return;
    }

    if (action === "save-settings") {
      const outputMode = data.outputMode as string | undefined;
      const sessionId = data.sessionId as string | undefined;
      const bypass = data.bypass as string | undefined;

      if (outputMode && (outputMode === "low" || outputMode === "medium" || outputMode === "high")) {
        if (sessionId) {
          this._sessionOutputModes.set(sessionId, outputMode as OutputMode);
        }
      }
      if (bypass !== undefined && sessionId) {
        const session = this.core.sessionManager.getSession(sessionId);
        if (session?.clientOverrides) {
          session.clientOverrides.bypassPermissions = bypass === "true";
        }
      }
      await sendText(context, "✅ Settings saved");
      return;
    }

    await sendText(context, `Unknown action: ${action}`);
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
          msteams: { width: "Full" },
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
          msteams: { width: "Full" },
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
   * Routes agent messages to Teams. The rate limiter handles per-conversation
   * serialization and throttling.
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

    try {
      await super.sendMessage(sessionId, content);
    } catch (err) {
      log.warn({ err, sessionId, type: content.type }, "[TeamsAdapter] Dispatch error");
    }
  }

  // ─── Handler overrides ───────────────────────────────────────────────────

  protected async handleThought(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {

    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const msg = this.composer.getOrCreate(sessionId, ctx.context);
    this.ensureSessionTitle(sessionId, msg);
    const summary = content.text?.split("\n")[0] || "";
    msg.addThinking(summary);
  }

  protected async handleText(sessionId: string, content: OutgoingMessage): Promise<void> {

    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const msg = this.composer.getOrCreate(sessionId, ctx.context);
    this.ensureSessionTitle(sessionId, msg);
    msg.closeActiveThinking();
    if (content.text) msg.addText(content.text);
  }

  protected async handleToolCall(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {

    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const msg = this.composer.getOrCreate(sessionId, ctx.context);
    this.ensureSessionTitle(sessionId, msg);
    msg.closeActiveThinking();
    const meta = (content.metadata ?? {}) as Partial<ToolCallMeta>;
    const toolName = meta.name || content.text || "Tool";
    const summary = formatToolSummary(
      toolName,
      meta.rawInput,
      meta.displaySummary as string | undefined,
    );
    const entryId = msg.addTimedStart("🔧", summary);
    this._toolEntryIds.set(sessionId, entryId);
  }

  protected async handleToolUpdate(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {

    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const msg = this.composer.getOrCreate(sessionId, ctx.context);
    const meta = (content.metadata ?? {}) as Partial<ToolCallMeta>;
    const toolName = meta.name || content.text || "";
    if (!toolName && !meta.displaySummary) return;
    const summary = formatToolSummary(
      toolName || "Tool",
      meta.rawInput,
      meta.displaySummary as string | undefined,
    );
    const entryId = this._toolEntryIds.get(sessionId);
    if (entryId) {
      msg.addTimedResult(entryId, summary);
      this._toolEntryIds.delete(sessionId);
    } else {
      msg.addTimedResult("", summary);
    }
  }

  /** Set the session title on the composer if not already set. */
  private ensureSessionTitle(sessionId: string, msg: import("./message-composer.js").SessionMessage): void {
    const session = this.core.sessionManager.getSession(sessionId);
    const name = session?.name;
    if (name) msg.setTitle(name);
  }

  protected async handlePlan(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {

    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const msg = this.composer.getOrCreate(sessionId, ctx.context);
    this.ensureSessionTitle(sessionId, msg);
    msg.closeActiveThinking();
    const planEntries = (content.metadata as { entries?: PlanEntry[] })?.entries ?? [];
    msg.setPlan(planEntries.map((e) => ({ content: e.content, status: e.status })));
  }

  protected async handleUsage(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const meta = content.metadata as { tokensUsed?: number; contextSize?: number; cost?: number; duration?: number } | undefined;
    if (meta?.tokensUsed == null) return; // No usage data — let handleSessionEnd finalize

    const msg = this.composer.getOrCreate(sessionId, ctx.context);
    const parts: string[] = [];
    parts.push(`${formatTokens(meta.tokensUsed)} tokens`);
    if (meta.duration != null) parts.push(`${(meta.duration / 1000).toFixed(1)}s`);
    if (meta.cost != null) parts.push(`$${meta.cost.toFixed(4)}`);
    parts.push("Done");
    msg.setUsage(parts.join(" · "));

    // Usage is the last event of a prompt turn — finalize the card and
    // clean up session state so the next turn gets a fresh start.
    await this.composer.finalize(sessionId);
    this.cleanupSessionState(sessionId);
  }

  /**
   * Clean up all per-session state (contexts, composer, output modes).
   */
  private cleanupSessionState(sessionId: string): void {
    const entry = this._sessionContexts.get(sessionId);
    const storedThreadId = entry?.threadId;
    if (storedThreadId && storedThreadId !== sessionId) {
      this._sessionContexts.delete(storedThreadId);
    }
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
    this._toolEntryIds.delete(sessionId);

    this.composer.cleanup(sessionId);
  }

  protected async handleSessionEnd(sessionId: string, _content: OutgoingMessage): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;

    const msg = this.composer.get(sessionId);
    if (msg) {
      msg.closeActiveThinking();
      const current = msg.getFooter();
      msg.setUsage(current ? `${current} · Task completed` : "Task completed");
    }
    await this.composer.finalize(sessionId);
    this.cleanupSessionState(sessionId);
  }

  protected async handleError(sessionId: string, content: OutgoingMessage): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const msg = this.composer.getOrCreate(sessionId, ctx.context);
    msg.closeActiveThinking();
    msg.addInfo("❌", "Error", content.text || "Unknown error");
    await this.composer.finalize(sessionId);
    this.cleanupSessionState(sessionId);
  }

  protected async handleAttachment(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (!content.attachment) return;
    const { attachment } = content;
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;

    // Strip TTS markers from body before sending attachment
    if (attachment.type === "audio") {
      const msg = this.composer.get(sessionId);
      if (msg) {
        await msg.stripPattern(/\[TTS\][\s\S]*?\[\/TTS\]/g).catch(() => {});
      }
    }

    if (isAttachmentTooLarge(attachment.size)) {
      log.warn({ sessionId, fileName: attachment.fileName, size: attachment.size }, "[TeamsAdapter] File too large");
      const msg = this.composer.getOrCreate(sessionId, ctx.context);
      msg.addResource(`📎 ⚠️ File too large (${Math.round(attachment.size / 1024 / 1024)}MB): ${attachment.fileName}`);
      return;
    }

    try {
      const shareUrl = await uploadFileViaGraph(
        this.graphClient,
        sessionId,
        attachment.filePath,
        attachment.fileName,
        attachment.mimeType,
      );

      const msg = this.composer.getOrCreate(sessionId, ctx.context);
      if (shareUrl) {
        msg.addResource(`📎 [${attachment.fileName}](${shareUrl})`);
      } else {
        msg.addResource(`📎 ${attachment.fileName} (${Math.round(attachment.size / 1024)}KB)`);
      }
    } catch (err) {
      log.error({ err, sessionId, fileName: attachment.fileName }, "[TeamsAdapter] Failed to send attachment");
    }
  }

  /**
   * Send an info entry. If a card is active (mid-turn), adds to it.
   * Otherwise sends a standalone info card that completes immediately.
   */
  private async sendInfo(sessionId: string, emoji: string, label: string, detail: string): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;

    if (this.composer.has(sessionId)) {
      // Mid-turn — add to the existing card
      const msg = this.composer.getOrCreate(sessionId, ctx.context);
      msg.addInfo(emoji, label, detail);
    } else {
      // Standalone — send a one-shot card (no "Working." animation)
      const card = TeamsAdapter.buildNotificationCard(emoji, label, detail);
      await sendCard(ctx.context, card);
    }
  }

  protected async handleSystem(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (!content.text) return;
    await this.sendInfo(sessionId, "⚙️", "System", content.text);
  }

  /** Sanitize metadata strings for safe markdown interpolation. */
  private static sanitizeMd(text: string): string {
    return text.replace(/[*_~`[\]()\\]/g, "").slice(0, 200);
  }

  protected async handleModeChange(sessionId: string, content: OutgoingMessage): Promise<void> {
    const modeId = TeamsAdapter.sanitizeMd(String((content.metadata as Record<string, unknown>)?.modeId ?? ""));
    await this.sendInfo(sessionId, "⚙️", "Mode", modeId);
  }

  protected async handleConfigUpdate(sessionId: string, content: OutgoingMessage): Promise<void> {
    const key = (content.metadata as Record<string, unknown>)?.key;
    const detail = key ? TeamsAdapter.sanitizeMd(String(key)) : "updated";
    await this.sendInfo(sessionId, "⚙️", "Config", detail);
  }

  protected async handleModelUpdate(sessionId: string, content: OutgoingMessage): Promise<void> {
    const modelId = TeamsAdapter.sanitizeMd(String((content.metadata as Record<string, unknown>)?.modelId ?? ""));
    await this.sendInfo(sessionId, "⚙️", "Model", modelId);
  }

  protected async handleUserReplay(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (!content.text) return;
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const msg = this.composer.getOrCreate(sessionId, ctx.context);
    msg.addText(content.text);
  }

  protected async handleResource(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (!content.text) return;
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const msg = this.composer.getOrCreate(sessionId, ctx.context);
    msg.addResource(`📎 ${content.text}`);
  }

  protected async handleResourceLink(sessionId: string, content: OutgoingMessage): Promise<void> {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) return;
    const rawUrl = (content.metadata as Record<string, unknown>)?.url as string | undefined;
    const rawName = (content.metadata as Record<string, unknown>)?.name as string | undefined;
    const url = rawUrl && /^https?:\/\//i.test(rawUrl) ? rawUrl : undefined;
    const name = rawName?.replace(/[\[\]\(\)]/g, "") || undefined;
    const text = url ? `📎 [${name || url}](${url})` : `📎 ${content.text || "Resource"}`;
    const msg = this.composer.getOrCreate(sessionId, ctx.context);
    msg.addResource(text);
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
    const typeLabel: Record<string, string> = {
      completed: "Completed", error: "Error", permission: "Permission", input_required: "Input Required", budget_warning: "Budget Warning",
    };

    const icon = typeIcon[notification.type] ?? "ℹ️";
    const label = typeLabel[notification.type] ?? "Notification";
    const detail = notification.sessionName
      ? `${notification.sessionName} — ${notification.summary}`
      : notification.summary;

    // Build a mini Adaptive Card matching the info Container style
    const card = TeamsAdapter.buildNotificationCard(icon, label, detail, notification.deepLink);

    // Post to the notification channel via Bot Framework REST API.
    if (this.notificationChannelId) {
      const ref = this.conversationStore.getAny();
      if (ref && TeamsAdapter.isValidServiceUrl(ref.serviceUrl)) {
        try {
          const botToken = await this.acquireBotToken();
          if (botToken) {
            const controller = new AbortController();
            const timeout = setTimeout(() => controller.abort(), 10_000);
            try {
              const response = await fetch(`${ref.serviceUrl}/v3/conversations/${encodeURIComponent(this.notificationChannelId)}/activities`, {
                method: "POST",
                headers: {
                  "Content-Type": "application/json",
                  "Authorization": `Bearer ${botToken}`,
                },
                body: JSON.stringify({
                  type: "message",
                  attachments: [{
                    contentType: "application/vnd.microsoft.card.adaptive",
                    content: card,
                  }],
                  from: { id: ref.botId, name: ref.botName },
                }),
                signal: controller.signal,
              });
              if (response.ok) return;
              log.warn({ status: response.status }, "[TeamsAdapter] Proactive notification to channel failed");
            } finally {
              clearTimeout(timeout);
            }
          }
        } catch (err) {
          log.warn({ err }, "[TeamsAdapter] Proactive notification error");
        }
      }
    }

    // Session-specific context fallback
    if (notification.sessionId) {
      const ctx = this._sessionContexts.get(notification.sessionId);
      if (ctx) {
        try {
          await sendCard(ctx.context, card);
          return;
        } catch (err) {
          log.debug({ err, sessionId: notification.sessionId }, "[TeamsAdapter] Session context fallback failed (context may be stale)");
        }
      }
    }

    log.debug({ type: notification.type, sessionName: notification.sessionName }, "[TeamsAdapter] sendNotification: no delivery path available");
  }

  /** Build a notification Adaptive Card with the same info Container style. */
  private static buildNotificationCard(emoji: string, label: string, detail: string, deepLink?: string): Record<string, unknown> {
    // Pre-escape detail, then append raw markdown link (buildLevel2 would double-escape it)
    const content = deepLink
      ? `${escapeMd(detail)}\n[Open →](${deepLink})`
      : escapeMd(detail);
    return {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "Container",
          spacing: "Small",
          items: [
            buildLevel1(emoji, escapeMd(label)),
            buildLevel2(content, undefined, true),
          ],
        },
      ],
      msteams: { width: "Full" },
    };
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

    // Update the main message title if the session has an active composer
    const msg = this.composer.get(sessionId);
    if (msg && newName) {
      msg.setTitle(newName);
    }

    log.debug({ sessionId, newName }, "[TeamsAdapter] renameSessionThread — title updated");
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