import { App } from "@microsoft/teams.apps";
import { BotBuilderPlugin } from "@microsoft/teams.botbuilder";
import { CardFactory, MemoryStorage } from "@microsoft/agents-hosting";
import type { TurnContext } from "@microsoft/agents-hosting";
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
import { TeamsDraftManager } from "./draft-manager.js";
import { PermissionHandler } from "./permissions.js";
import { handleCommand, setupCardActionCallbacks } from "./commands/index.js";
import { spawnAssistant, buildWelcomeMessage } from "./assistant.js";
import { downloadTeamsFile, isAttachmentTooLarge, buildFileAttachmentCard } from "./media.js";
import type { OutputMode } from "./activity.js";

export class TeamsAdapter extends MessagingAdapter {
  readonly name = "teams";
  readonly renderer: IRenderer = new TeamsRenderer();
  readonly capabilities: AdapterCapabilities = {
    streaming: true,
    richFormatting: true,
    threads: true,
    reactions: false,
    fileUpload: true,
    voice: true,
  };

  readonly core: OpenACPCore;
  private app: App;
  private teamsConfig: TeamsChannelConfig;
  private sendQueue: SendQueue;
  private draftManager: TeamsDraftManager;
  private permissionHandler!: PermissionHandler;

  private notificationChannelId?: string;
  private notificationConversationId?: string;
  private notificationContext?: TurnContext;
  private assistantSession: Session | null = null;
  private assistantInitializing = false;
  private fileService: FileServiceInterface;

  private _sessionContexts = new Map<string, { context: TurnContext; isAssistant: boolean }>();
  private _sessionOutputModes = new Map<string, OutputMode>();

  constructor(core: OpenACPCore, config: TeamsChannelConfig) {
    super(
      { configManager: core.configManager },
      { ...config as unknown as Record<string, unknown>, maxMessageLength: 2000, enabled: config.enabled ?? true } as MessagingAdapterConfig,
    );
    this.core = core;
    this.teamsConfig = config;
    this.sendQueue = new SendQueue({ minInterval: 1000 });
    this.draftManager = new TeamsDraftManager(this.sendQueue as any);
    this.fileService = core.fileService;

    const botBuilderPlugin = new BotBuilderPlugin();
    this.app = new App({
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

      await this.app.start();

      log.info("[TeamsAdapter] Initialization complete");
    } catch (err) {
      log.error({ err }, "[TeamsAdapter] Initialization failed");
      throw err;
    }
  }

  // ─── stop ─────────────────────────────────────────────────────────────────

  async stop(): Promise<void> {
    if (this.assistantSession) {
      try {
        await this.assistantSession.destroy();
      } catch (err) {
        log.warn({ err }, "[TeamsAdapter] Failed to destroy assistant session");
      }
      this.assistantSession = null;
    }
    this.notificationContext = undefined;
    this._sessionContexts.clear();
    this._sessionOutputModes.clear();
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
        await context.sendActivity({ text: `✅ File consent accepted: ${fileName}` } as any);
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
        await context.sendActivity({ text: `❌ File consent declined: ${fileName}` } as any);
        return { status: 200 };
      } catch (err) {
        log.error({ err }, "[TeamsAdapter] file.consent.decline error");
        return { status: 500 };
      }
    });

    this.app.on("message", async (context: any) => {
      const text = context.activity.text ?? "";
      const userId = context.activity.from?.id ?? "unknown";
      const channelId = String(context.activity.conversation?.id ?? "unknown");
      const threadId = String(context.activity.channelId ?? channelId);

      try {
        log.debug({ threadId, userId, text: text.slice(0, 50) }, "[TeamsAdapter] message received");

        if (!text && !context.activity.attachments?.length) return;

        const sessionId =
          this.core.sessionManager.getSessionByThread("teams", threadId)?.id ?? "unknown";

        // Capture notification channel context on first message
        if (this.notificationChannelId && threadId === this.notificationChannelId) {
          this.notificationConversationId = channelId;
          this.notificationContext = context;
        }

        const attachments: Attachment[] = [];
        if (context.activity.attachments?.length) {
          for (const att of context.activity.attachments) {
            try {
              const buffer = await downloadTeamsFile(att.contentUrl ?? "", att.name ?? "attachment");
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
          // Reset draft state for new prompt
          const draft = this.draftManager.getDraft(sessionId);
          if (draft) {
            this.draftManager.cleanup(sessionId);
          }
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
          await context.sendActivity({ text: "❌ Failed to process message. Please try again." } as any);
        } catch { /* best effort */ }
      }
    });
  }

  // ─── Card action handler ────────────────────────────────────────────────

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

      if (data?.sessionId && data?.callbackKey && data?.requestId) {
        const handled = await this.permissionHandler.handleCardAction(
          context,
          verb,
          data.sessionId as string,
          data.callbackKey as string,
          data.requestId as string,
        );
        if (handled) return;
      }

      if (typeof verb === "string" && verb.startsWith("om:")) {
        const parts = verb.split(":");
        if (parts.length === 3) {
          const [_prefix, sessionId, mode] = parts;
          if (mode === "low" || mode === "medium" || mode === "high") {
            this._sessionOutputModes.set(sessionId, mode as OutputMode);
            await context.sendActivity({ text: `🔄 Output mode: **${mode}**` } as any);
            return;
          }
        }
      }

      if (verb.startsWith("cancel:")) {
        const sessionId = verb.split(":")[1];
        if (sessionId) {
          const session = this.core.sessionManager.getSession(sessionId);
          if (session) {
            try {
              await session.destroy();
            } catch (err) {
              log.error({ err, sessionId }, "[TeamsAdapter] cancel: destroy failed");
              try {
                await context.sendActivity({ text: "❌ Failed to cancel session" } as any);
              } catch { /* best effort */ }
              return;
            }
            try {
              await context.sendActivity({ text: "❌ Session cancelled" } as any);
            } catch { /* best effort */ }
          }
        }
        return;
      }

      await setupCardActionCallbacks(context, this);
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
            await (context.sendActivity as any)({ text: content });
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
      await context.sendActivity({ text } as any);
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
          version: "1.4",
          body: [
            { type: "TextBlock", text: response.title, weight: "Bolder" as const, size: "Medium" as const },
            { type: "TextBlock", text: response.options.map((o) => `• ${o.label}`).join("\n"), wrap: true },
          ],
          actions: response.options.slice(0, 5).map((opt) => ({
            type: "Action.Execute" as const,
            title: opt.label.slice(0, 20),
            data: { verb: `cmd:${opt.command}` },
          })),
        };
        await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] } as any);
        break;
      }
      case "list": {
        const text = response.items.map((i) => `• **${i.label}**${i.detail ? ` — ${i.detail}` : ""}`).join("\n");
        await reply(`${response.title}\n${text}`);
        break;
      }
      case "confirm": {
        const card = {
          type: "AdaptiveCard",
          version: "1.4",
          body: [{ type: "TextBlock", text: response.question, wrap: true }],
          actions: [
            { type: "Action.Execute" as const, title: "Yes", data: { verb: `cmd:${response.onYes}` } },
            { type: "Action.Execute" as const, title: "No", data: { verb: "cmd:noop" } },
          ],
        };
        await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] } as any);
        break;
      }
      case "silent":
        break;
      default:
        await reply(`⚠️ Unexpected response type: ${(response as any).type ?? "unknown"}`);
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
      ready.finally(() => {
        this.assistantInitializing = false;
      });
    } catch (err) {
      this.assistantInitializing = false;
      log.error({ err }, "[TeamsAdapter] Failed to spawn assistant");
    }
  }

  async respawnAssistant(): Promise<void> {
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

  // ─── Helper: resolve context ─────────────────────────────────────────────

  private async getContext(sessionId: string): Promise<TurnContext | null> {
    const ctx = this._sessionContexts.get(sessionId);
    return ctx?.context ?? null;
  }

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

  private getSessionContext(sessionId: string): { context: TurnContext; isAssistant: boolean } {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) {
      throw new Error(`No context stored for session ${sessionId}`);
    }
    return ctx;
  }

  // ─── sendMessage ─────────────────────────────────────────────────────────

  async sendMessage(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (
      this.assistantInitializing &&
      this.assistantSession &&
      sessionId === this.assistantSession.id
    ) {
      return;
    }

    const context = await this.getContext(sessionId);
    if (!context) return;

    const isAssistant = this.assistantSession != null && sessionId === this.assistantSession.id;

    this._sessionContexts.set(sessionId, { context, isAssistant });

    try {
      await super.sendMessage(sessionId, content);
    } finally {
      // Give nested async handlers time to complete before deleting context
      await new Promise((resolve) => setTimeout(resolve, 0));
      this._sessionContexts.delete(sessionId);
    }
  }

  // ─── Handler overrides ───────────────────────────────────────────────────

  protected async handleThought(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    // Thoughts are not sent as messages in Teams — buffered and displayed via plan
    void sessionId;
    void content;
  }

  protected async handleText(sessionId: string, content: OutgoingMessage): Promise<void> {
    let ctx: { context: TurnContext; isAssistant: boolean } | null = null;
    try {
      ctx = this.getSessionContext(sessionId);
    } catch (err) {
      log.warn({ err, sessionId }, "[TeamsAdapter] handleText: no session context");
      return;
    }
    const { context } = ctx;
    const draft = this.draftManager.getOrCreate(sessionId, context);
    draft.append(content.text);
    this.draftManager.appendText(sessionId, content.text);
  }

  protected async handleToolCall(sessionId: string, content: OutgoingMessage, verbosity: DisplayVerbosity): Promise<void> {
    let ctx: { context: TurnContext; isAssistant: boolean } | null = null;
    try {
      ctx = this.getSessionContext(sessionId);
    } catch (err) {
      log.warn({ err, sessionId }, "[TeamsAdapter] handleToolCall: no session context");
      return;
    }
    const { context, isAssistant } = ctx;
    await this.draftManager.finalize(sessionId, context, isAssistant);
    try {
      const rendered = this.renderer.renderToolCall(content, verbosity);
      await context.sendActivity({ text: rendered.body } as any);
    } catch (err) {
      log.error({ err, sessionId }, "[TeamsAdapter] handleToolCall: sendActivity failed");
    }
  }

  protected async handleToolUpdate(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    // Tool updates rendered in plan/usage cards — no separate handling needed
    void sessionId;
    void content;
  }

  protected async handlePlan(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    let ctx: { context: TurnContext; isAssistant: boolean } | null = null;
    try {
      ctx = this.getSessionContext(sessionId);
    } catch (err) {
      log.warn({ err, sessionId }, "[TeamsAdapter] handlePlan: no session context");
      return;
    }
    const { context } = ctx;
    const entries = (content.metadata as { entries?: PlanEntry[] })?.entries ?? [];
    const mode = this.resolveMode(sessionId);
    const { formatPlan } = await import("./formatting.js");
    const text = formatPlan(entries, mode);
    const card = {
      type: "AdaptiveCard" as const,
      version: "1.4" as const,
      body: [{ type: "TextBlock" as const, text, wrap: true }],
    };
    try {
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card as any)] } as any);
    } catch (err) {
      log.error({ err, sessionId }, "[TeamsAdapter] handlePlan: sendActivity failed");
    }
  }

  protected async handleUsage(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    let ctx: { context: TurnContext; isAssistant: boolean } | null = null;
    try {
      ctx = this.getSessionContext(sessionId);
    } catch (err) {
      log.warn({ err, sessionId }, "[TeamsAdapter] handleUsage: no session context");
      return;
    }
    const { context, isAssistant } = ctx;
    await this.draftManager.finalize(sessionId, context, isAssistant);
    const meta = content.metadata as { tokensUsed?: number; contextSize?: number; cost?: number; duration?: number } | undefined;
    const mode = this.resolveMode(sessionId);
    const { renderUsageCard } = await import("./formatting.js");
    const { body } = renderUsageCard(meta ?? {}, mode);
    const card = { type: "AdaptiveCard" as const, version: "1.4" as const, body };
    try {
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card as any)] } as any);
    } catch (err) {
      log.error({ err, sessionId }, "[TeamsAdapter] handleUsage: sendActivity failed");
    }

    if (this.notificationChannelId && sessionId !== this.assistantSession?.id) {
      const sess = this.core.sessionManager.getSession(sessionId);
      const name = sess?.name || "Session";
      try {
        await context.sendActivity({ text: `✅ **${name}** — Task completed.` } as any);
      } catch { /* best effort */ }
    }
  }

  protected async handleSessionEnd(sessionId: string, _content: OutgoingMessage): Promise<void> {
    let ctx: { context: TurnContext; isAssistant: boolean } | null = null;
    try {
      ctx = this.getSessionContext(sessionId);
    } catch (err) {
      log.warn({ err, sessionId }, "[TeamsAdapter] handleSessionEnd: no session context");
      return;
    }
    const { context, isAssistant } = ctx;
    await this.draftManager.finalize(sessionId, context, isAssistant);
    this.draftManager.cleanup(sessionId);
    this._sessionOutputModes.delete(sessionId);
    try {
      await context.sendActivity({ text: "✅ **Done**" } as any);
    } catch { /* best effort */ }
  }

  protected async handleError(sessionId: string, content: OutgoingMessage): Promise<void> {
    let ctx: { context: TurnContext; isAssistant: boolean } | null = null;
    try {
      ctx = this.getSessionContext(sessionId);
    } catch (err) {
      log.warn({ err, sessionId }, "[TeamsAdapter] handleError: no session context");
      return;
    }
    const { context, isAssistant } = ctx;
    await this.draftManager.finalize(sessionId, context, isAssistant);
    this.draftManager.cleanup(sessionId);
    try {
      await context.sendActivity({ text: `❌ **Error:** ${content.text}` } as any);
    } catch { /* best effort */ }
  }

  protected async handleAttachment(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (!content.attachment) return;
    const { attachment } = content;
    let ctx: { context: TurnContext; isAssistant: boolean } | null = null;
    try {
      ctx = this.getSessionContext(sessionId);
    } catch (err) {
      log.warn({ err, sessionId }, "[TeamsAdapter] handleAttachment: no session context");
      return;
    }
    const { context, isAssistant } = ctx;
    await this.draftManager.finalize(sessionId, context, isAssistant);

    if (isAttachmentTooLarge(attachment.size)) {
      log.warn({ sessionId, fileName: attachment.fileName, size: attachment.size }, "[TeamsAdapter] File too large");
      try {
        await context.sendActivity({
          text: `⚠️ File too large to send (${Math.round(attachment.size / 1024 / 1024)}MB): ${attachment.fileName}`,
        } as any);
      } catch { /* best effort */ }
      return;
    }

    try {
      // Show file card — actual serving requires Graph API upload to SharePoint/OneDrive
      const card = buildFileAttachmentCard(
        attachment.fileName,
        attachment.size,
        `file://${attachment.filePath}`,
        attachment.mimeType,
      );
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card as any)] } as any);

      if (attachment.type === "audio") {
        const draft = this.draftManager.getDraft(sessionId);
        if (draft) {
          draft.stripPattern(/\[TTS\][\s\S]*?\[\/TTS\]/g).catch((err) => {
            log.warn({ err, sessionId }, "[TeamsAdapter] handleAttachment: stripPattern failed");
          });
        }
      }
    } catch (err) {
      log.error({ err, sessionId, fileName: attachment.fileName }, "[TeamsAdapter] Failed to send attachment");
    }
  }

  protected async handleSystem(sessionId: string, content: OutgoingMessage): Promise<void> {
    let ctx: { context: TurnContext; isAssistant: boolean } | null = null;
    try {
      ctx = this.getSessionContext(sessionId);
    } catch (err) {
      log.warn({ err, sessionId }, "[TeamsAdapter] handleSystem: no session context");
      return;
    }
    const { context } = ctx;
    try {
      await context.sendActivity({ text: content.text } as any);
    } catch { /* best effort */ }
  }

  // ─── sendPermissionRequest ──────────────────────────────────────────────

  async sendPermissionRequest(sessionId: string, request: PermissionRequest): Promise<void> {
    const session = this.core.sessionManager.getSession(sessionId);
    if (!session) {
      log.warn({ sessionId }, "[TeamsAdapter] sendPermissionRequest: session not found");
      return;
    }
    const context = await this.getContext(sessionId);
    if (!context) return;
    await this.permissionHandler.sendPermissionRequest(session, request, context);
  }

  // ─── sendNotification ──────────────────────────────────────────────────

  async sendNotification(notification: NotificationMessage): Promise<void> {
    const typeIcon: Record<string, string> = {
      completed: "✅", error: "❌", permission: "🔐", input_required: "💬",
    };

    const icon = typeIcon[notification.type] ?? "ℹ️";
    const name = notification.sessionName ? ` **${notification.sessionName}**` : "";
    let text = `${icon}${name}: ${notification.summary}`;
    if (notification.deepLink) {
      text += `\n${notification.deepLink}`;
    }

    // Route to notification thread if we have context for it
    if (this.notificationContext && this.notificationChannelId) {
      try {
        await this.notificationContext.sendActivity(text as any);
        return;
      } catch (err) {
        log.warn({ err, type: notification.type, sessionName: notification.sessionName }, "[TeamsAdapter] Failed to send notification to channel");
      }
    }

    log.debug({ type: notification.type, sessionName: notification.sessionName, text }, "[TeamsAdapter] sendNotification: no notificationContext configured, skipping");
  }

  // ─── createSessionThread ─────────────────────────────────────────────────

  async createSessionThread(sessionId: string, name: string): Promise<string> {
    const threadId = `thread-${sessionId}-${Date.now()}`;
    const session = this.core.sessionManager.getSession(sessionId);
    if (session) {
      session.threadId = threadId;
    }
    const record = this.core.sessionManager.getSessionRecord(sessionId);
    if (record) {
      await this.core.sessionManager.patchRecord(sessionId, {
        platform: { ...record.platform, threadId },
      });
    }
    return threadId;
  }

  async renameSessionThread(sessionId: string, newName: string): Promise<void> {
    const record = this.core.sessionManager.getSessionRecord(sessionId);
    if (!record) return;
    const threadId = record.platform?.threadId as string | undefined;
    if (!threadId) return;

    // Teams Graph API would be needed for actual rename:
    // PATCH https://graph.microsoft.com/v1.0/chats/{chat-id}
    // Requires Chat.ReadWrite.All permission
    log.info({ sessionId, threadId, newName }, "[TeamsAdapter] renameSessionThread — requires Graph API (Chat.ReadWrite.All)");
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