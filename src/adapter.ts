import { App, MemoryStorage } from "@microsoft/teams.apps";
import { configureBotBuilderAdapter } from "@microsoft/teams.botbuilder";
import { CardFactory } from "@microsoft/teams.cards";
import type { TurnContext } from "@microsoft/teams.botbuilder";
import type {
  OutgoingMessage,
  PermissionRequest,
  NotificationMessage,
  AgentCommand,
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
import { ActivityTracker, type ToolCallMeta, type OutputMode } from "./activity.js";
import { PermissionHandler } from "./permissions.js";
import {
  handleCommand,
  setupCardActionCallbacks,
} from "./commands/index.js";
import { spawnAssistant, buildWelcomeMessage } from "./assistant.js";
import {
  downloadTeamsFile,
  isAttachmentTooLarge,
} from "./media.js";

export class TeamsAdapter extends MessagingAdapter {
  readonly name = "teams";
  readonly renderer: IRenderer = new TeamsRenderer();
  readonly capabilities: AdapterCapabilities = {
    streaming: true,
    richFormatting: true,
    threads: true,
    reactions: false, // Teams doesn't have reactions like Discord
    fileUpload: true,
    voice: true,
  };

  readonly core: OpenACPCore;
  private app: App;
  private teamsConfig: TeamsChannelConfig;
  private sendQueue: SendQueue;
  private draftManager: TeamsDraftManager;
  private _outputModeResolver = new OutputModeResolver();
  private permissionHandler!: PermissionHandler;
  private sessionTrackers: Map<string, ActivityTracker> = new Map();

  private notificationChannelId?: string;
  private assistantSession: Session | null = null;
  private assistantInitializing = false;
  private fileService: FileServiceInterface;

  // Per-session context for concurrency safety in sendMessage handlers
  private _sessionContexts = new Map<string, { context: TurnContext; isAssistant: boolean }>();

  constructor(core: OpenACPCore, config: TeamsChannelConfig) {
    super(
      { configManager: core.configManager },
      { ...config as Record<string, unknown>, maxMessageLength: 2000, enabled: config.enabled ?? true } as MessagingAdapterConfig,
    );
    this.core = core;
    this.teamsConfig = config;
    this.sendQueue = new SendQueue({ minInterval: 1000 });
    this.draftManager = new TeamsDraftManager(this.sendQueue);
    this.fileService = core.fileService;

    const adapter = configureBotBuilderAdapter({
      appId: config.botAppId,
      appPassword: config.botAppPassword,
    });

    this.app = new App({
      storage: new MemoryStorage(),
      adapter,
    });
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
    await this.app.stop();
    log.info("[TeamsAdapter] Stopped");
  }

  // ─── Message handler ──────────────────────────────────────────────────────

  private setupMessageHandler(): void {
    this.app.on("message", async (context) => {
      try {
        const text = context.activity.text ?? "";
        const userId = context.activity.from?.id ?? "unknown";
        const channelId = context.activity.conversation?.id ?? "unknown";
        const threadId = context.activity.channelId ?? channelId;

        log.debug(
          { threadId, userId, text: text.slice(0, 50) },
          "[TeamsAdapter] message received",
        );

        // Ignore empty messages without attachments
        if (!text && !context.activity.attachments?.length) return;

        // Resolve sessionId for file storage (fallback to "unknown" for new sessions)
        const sessionId =
          this.core.sessionManager.getSessionByThread("teams", threadId)
            ?.id ?? "unknown";

        // Process attachments
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

        // Generate fallback text if message has attachments but no text
        let messageText = text;
        if (!messageText && attachments.length > 0) {
          messageText = attachments.map((a) => `[Attachment: ${a.fileName}]`).join("\n");
        }

        if (!messageText && attachments.length === 0) return;

        // Route assistant thread messages to assistant
        if (
          this.teamsConfig.assistantThreadId &&
          threadId === this.teamsConfig.assistantThreadId
        ) {
          if (this.assistantSession && messageText) {
            await this.assistantSession.enqueuePrompt(
              messageText,
              attachments.length > 0 ? attachments : undefined,
            );
          }
          return;
        }

        // Reset tracker state for new prompt cycle on existing sessions
        if (sessionId !== "unknown") {
          const tracker = this.sessionTrackers.get(sessionId);
          if (tracker) {
            await tracker.onNewPrompt();
          }
        }

        // Route to core for session dispatch
        await this.core.handleMessage({
          channelId: "teams",
          threadId,
          userId,
          text: messageText,
          ...(attachments.length > 0 ? { attachments } : {}),
        });
      } catch (err) {
        log.error({ err }, "[TeamsAdapter] message handler error");
      }
    });
  }

  // ─── Card action handler (Adaptive Card invoke) ─────────────────────────

  private setupCardActionHandler(): void {
    this.app.on("adaptiveCardInvoke", async (context) => {
      try {
        const action = context.activity.value?.action;
        if (!action) return;

        const verb = action.verb;
        const data = action.data;

        if (!verb) return;

        // Permission button callbacks
        if (data?.sessionId && data?.callbackKey && data?.requestId) {
          const handled = await this.permissionHandler.handleCardAction(
            context,
            verb,
            data.sessionId,
            data.callbackKey,
            data.requestId,
          );
          if (handled) return;
        }

        // Output mode buttons (om:sessionId:mode)
        if (typeof verb === "string" && verb.startsWith("om:")) {
          const parts = verb.split(":");
          if (parts.length === 3) {
            const [_prefix, sessionId, mode] = parts;
            if (mode === "low" || mode === "medium" || mode === "high") {
              this.updateSessionOutputMode(sessionId, mode as OutputMode);
              await context.sendActivity({ text: `🔄 Output mode: **${mode}**` });
              return;
            }
          }
        }

        // Cancel button
        if (verb.startsWith("cancel:")) {
          const sessionId = verb.split(":")[1];
          if (sessionId) {
            const session = this.core.sessionManager.getSession(sessionId);
            if (session) {
              await session.destroy();
              await context.sendActivity({ text: "❌ Session cancelled" });
            }
          }
          return;
        }

        // Route to card action callbacks (command buttons)
        await setupCardActionCallbacks(context, this);
      } catch (err) {
        log.error({ err }, "[TeamsAdapter] adaptiveCardInvoke handler error");
      }
    });
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

    const channelId = context.activity.conversation?.id ?? "teams";

    const response = await registry.execute(text, {
      raw: "",
      sessionId,
      channelId: "teams",
      userId,
      reply: async (content: string) => {
        if (typeof content === "string") {
          await context.sendActivity({ text: content });
        }
      },
    });

    if (response.type !== "silent") {
      await this.renderCommandResponse(response, context);
    }
  }

  private async renderCommandResponse(
    response: CommandResponse,
    context: TurnContext,
  ): Promise<void> {
    const reply = async (text: string) => {
      await context.sendActivity({ text });
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
          type: "AdaptiveCard" as const,
          version: "1.4" as const,
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
        await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
        break;
      }
      case "list": {
        const text = response.items.map((i) => `• **${i.label}**${i.detail ? ` — ${i.detail}` : ""}`).join("\n");
        await reply(`${response.title}\n${text}`);
        break;
      }
      case "confirm": {
        const card = {
          type: "AdaptiveCard" as const,
          version: "1.4" as const,
          body: [
            { type: "TextBlock", text: response.question, wrap: true },
          ],
          actions: [
            { type: "Action.Execute" as const, title: "Yes", data: { verb: `cmd:${response.onYes}` } },
            { type: "Action.Execute" as const, title: "No", data: { verb: "cmd:noop" } },
          ],
        };
        await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
        break;
      }
      case "silent":
        break;
    }
  }

  // ─── Assistant ────────────────────────────────────────────────────────────

  private async setupAssistant(): Promise<void> {
    let threadId = this.teamsConfig.assistantThreadId ?? undefined;

    if (!threadId) {
      // Create a new thread for the assistant
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

  // ─── Helper: resolve thread ────────────────────────────────────────────────

  private async getContext(sessionId: string): Promise<TurnContext | null> {
    const ctx = this._sessionContexts.get(sessionId);
    return ctx?.context ?? null;
  }

  // ─── Helper: get or create activity tracker ──────────────────────────────

  private resolveMode(sessionId: string): OutputMode {
    return this._outputModeResolver.resolve(
      this.core.configManager as any,
      this.name,
      sessionId,
      this.core.sessionManager as any,
    );
  }

  private getOrCreateTracker(
    sessionId: string,
    context: TurnContext,
    outputMode: OutputMode = "medium",
  ): ActivityTracker {
    let tracker = this.sessionTrackers.get(sessionId);
    if (!tracker) {
      tracker = new ActivityTracker(
        context,
        this.sendQueue,
        outputMode,
        sessionId,
      );
      this.sessionTrackers.set(sessionId, tracker);
    } else {
      tracker.setOutputMode(outputMode);
    }
    return tracker;
  }

  updateSessionOutputMode(sessionId: string, mode: OutputMode): void {
    const tracker = this.sessionTrackers.get(sessionId);
    if (!tracker) return;
    tracker.setOutputMode(mode);
    tracker.rerender();
  }

  private getSessionContext(sessionId: string): { context: TurnContext; isAssistant: boolean } {
    const ctx = this._sessionContexts.get(sessionId);
    if (!ctx) {
      throw new Error(`No context stored for session ${sessionId}`);
    }
    return ctx;
  }

  // ─── sendMessage ────────────────────────────────────────────────────────────

  async sendMessage(
    sessionId: string,
    content: OutgoingMessage,
  ): Promise<void> {
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
      this._sessionContexts.delete(sessionId);
    }
  }

  // ─── Handler overrides ─────────────────────────────────────────────────────

  protected async handleThought(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    const { context } = this.getSessionContext(sessionId);
    const mode = this.resolveMode(sessionId);
    const tracker = this.getOrCreateTracker(sessionId, context, mode);
    await tracker.onThought(content.text || "");
  }

  protected async handleText(sessionId: string, content: OutgoingMessage): Promise<void> {
    const { context } = this.getSessionContext(sessionId);
    if (!this.draftManager.hasDraft(sessionId)) {
      const mode = this.resolveMode(sessionId);
      const tracker = this.getOrCreateTracker(sessionId, context, mode);
      await tracker.onTextStart();
    }
    const draft = this.draftManager.getOrCreate(sessionId, context);
    draft.append(content.text);
    this.draftManager.appendText(sessionId, content.text);
  }

  protected async handleToolCall(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    const { context, isAssistant } = this.getSessionContext(sessionId);
    const meta = (content.metadata ?? {}) as Partial<ToolCallMeta>;
    const mode = this.resolveMode(sessionId);
    const tracker = this.getOrCreateTracker(sessionId, context, mode);
    await this.draftManager.finalize(sessionId, context, isAssistant);
    await tracker.onToolCall(
      {
        id: meta.id ?? "",
        name: meta.name ?? content.text ?? "Tool",
        kind: meta.kind,
        status: meta.status,
        content: meta.content,
        rawInput: meta.rawInput,
        viewerLinks: meta.viewerLinks,
        viewerFilePath: meta.viewerFilePath,
        displaySummary: meta.displaySummary as string | undefined,
        displayTitle: meta.displayTitle as string | undefined,
        displayKind: meta.displayKind as string | undefined,
      },
      String(meta.kind ?? ""),
      meta.rawInput,
    );
  }

  protected async handleToolUpdate(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    const { context } = this.getSessionContext(sessionId);
    const meta = (content.metadata ?? {}) as Partial<ToolCallMeta & { diffStats?: { added: number; removed: number } }>;
    const mode = this.resolveMode(sessionId);
    const tracker = this.getOrCreateTracker(sessionId, context, mode);
    await tracker.onToolUpdate(
      meta.id ?? "",
      meta.status ?? "completed",
      meta.viewerLinks as { file?: string; diff?: string } | undefined,
      typeof meta.content === "string" ? meta.content : null,
      meta.rawInput ?? undefined,
      meta.diffStats as { added: number; removed: number } | undefined,
    );
  }

  protected async handlePlan(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    const { context } = this.getSessionContext(sessionId);
    const meta = (content.metadata ?? {}) as { entries?: PlanEntry[] };
    const entries = meta.entries ?? [];
    const mode = this.resolveMode(sessionId);
    const tracker = this.getOrCreateTracker(sessionId, context, mode);
    await tracker.onPlan(entries);
  }

  protected async handleUsage(sessionId: string, content: OutgoingMessage, _verbosity: DisplayVerbosity): Promise<void> {
    const { context, isAssistant } = this.getSessionContext(sessionId);
    await this.draftManager.finalize(sessionId, context, isAssistant);
    const meta = content.metadata as { tokensUsed?: number; contextSize?: number; cost?: number; duration?: number } | undefined;
    const mode = this.resolveMode(sessionId);

    try {
      const { renderUsageCard } = await import("./formatting.js");
      const { body } = renderUsageCard(meta ?? {}, mode);
      const card = {
        type: "AdaptiveCard" as const,
        version: "1.4" as const,
        body,
      };
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
    } catch (err) {
      log.warn({ err, sessionId }, "Failed to send usage card");
    }

    // Notify notification channel
    if (this.notificationChannelId && sessionId !== this.assistantSession?.id) {
      const sess = this.core.sessionManager.getSession(sessionId);
      const name = sess?.name || "Session";
      try {
        await this.core.lifecycleManager?.serviceRegistry?.get("adapter:teams");
        // TODO: Send to notification channel
      } catch { /* best effort */ }
    }
  }

  protected async handleSessionEnd(sessionId: string, _content: OutgoingMessage): Promise<void> {
    const { context, isAssistant } = this.getSessionContext(sessionId);
    await this.draftManager.finalize(sessionId, context, isAssistant);
    this.draftManager.cleanup(sessionId);
    const tracker = this.sessionTrackers.get(sessionId);
    if (tracker) {
      await tracker.cleanup();
      this.sessionTrackers.delete(sessionId);
    } else {
      try {
        await context.sendActivity({ text: "✅ **Done**" });
      } catch { /* best effort */ }
    }
  }

  protected async handleError(sessionId: string, content: OutgoingMessage): Promise<void> {
    const { context, isAssistant } = this.getSessionContext(sessionId);
    await this.draftManager.finalize(sessionId, context, isAssistant);
    const tracker = this.sessionTrackers.get(sessionId);
    if (tracker) {
      tracker.destroy();
      this.sessionTrackers.delete(sessionId);
    }
    try {
      await context.sendActivity({ text: `❌ **Error:** ${content.text}` });
    } catch { /* best effort */ }
  }

  protected async handleAttachment(sessionId: string, content: OutgoingMessage): Promise<void> {
    if (!content.attachment) return;
    const { attachment } = content;
    const { context, isAssistant } = this.getSessionContext(sessionId);
    await this.draftManager.finalize(sessionId, context, isAssistant);

    if (isAttachmentTooLarge(attachment.size)) {
      log.warn({ sessionId, fileName: attachment.fileName, size: attachment.size }, "[TeamsAdapter] File too large");
      try {
        await context.sendActivity({
          text: `⚠️ File too large to send (${Math.round(attachment.size / 1024 / 1024)}MB): ${attachment.fileName}`,
        });
      } catch { /* best effort */ }
      return;
    }

    try {
      // Teams doesn't support direct file send like Discord
      // Send a message with the file name reference
      await context.sendActivity({
        text: `📎 Attachment: ${attachment.fileName}`,
      });

      if (attachment.type === "audio") {
        const draft = this.draftManager.getDraft(sessionId);
        if (draft) {
          draft.stripPattern(/\[TTS\][\s\S]*?\[\/TTS\]/g).catch(() => {});
        }
      }
    } catch (err) {
      log.error({ err, sessionId, fileName: attachment.fileName }, "[TeamsAdapter] Failed to send attachment");
    }
  }

  protected async handleSystem(sessionId: string, content: OutgoingMessage): Promise<void> {
    const { context } = this.getSessionContext(sessionId);
    try {
      await context.sendActivity({ text: content.text });
    } catch { /* best effort */ }
  }

  // ─── sendPermissionRequest ────────────────────────────────────────────────

  async sendPermissionRequest(
    sessionId: string,
    request: PermissionRequest,
  ): Promise<void> {
    const session = this.core.sessionManager.getSession(sessionId);
    if (!session) {
      log.warn({ sessionId }, "[TeamsAdapter] sendPermissionRequest: session not found");
      return;
    }

    const context = await this.getContext(sessionId);
    if (!context) return;

    await this.permissionHandler.sendPermissionRequest(session, request, context);
  }

  // ─── sendNotification ─────────────────────────────────────────────────────

  async sendNotification(notification: NotificationMessage): Promise<void> {
    if (!this.notificationChannelId) return;

    const typeIcon: Record<string, string> = {
      completed: "✅",
      error: "❌",
      permission: "🔐",
      input_required: "💬",
    };

    const icon = typeIcon[notification.type] ?? "ℹ️";
    const name = notification.sessionName ? ` **${notification.sessionName}**` : "";
    let text = `${icon}${name}: ${notification.summary}`;
    if (notification.deepLink) {
      text += `\n${notification.deepLink}`;
    }

    try {
      // TODO: Send to notification channel via Teams API
      log.info({ text }, "[TeamsAdapter] Notification (not yet sent to channel)");
    } catch (err) {
      log.warn({ err }, "[TeamsAdapter] Failed to send notification");
    }
  }

  // ─── createSessionThread ─────────────────────────────────────────────────

  async createSessionThread(sessionId: string, name: string): Promise<string> {
    // In Teams, threads are created within channels
    // This would use Teams SDK to create a new reply/thread in the channel
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

  // ─── renameSessionThread ──────────────────────────────────────────────────

  async renameSessionThread(sessionId: string, newName: string): Promise<void> {
    const session = this.core.sessionManager.getSession(sessionId);
    const threadId = session?.threadId;
    if (!threadId) return;
    // TODO: Implement Teams thread rename via SDK
    log.info({ sessionId, threadId, newName }, "[TeamsAdapter] renameSessionThread not yet implemented");
  }

  // ─── deleteSessionThread ──────────────────────────────────────────────────

  async deleteSessionThread(sessionId: string): Promise<void> {
    const session = this.core.sessionManager.getSession(sessionId);
    const threadId = session?.threadId;
    if (!threadId) return;
    // TODO: Implement Teams thread deletion via SDK
    log.info({ sessionId, threadId }, "[TeamsAdapter] deleteSessionThread not yet implemented");
  }

  // ─── Public helpers (for slash commands) ─────────────────────────────────

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
}

// ─── OutputModeResolver ────────────────────────────────────────────────────────

class OutputModeResolver {
  resolve(
    configManager: { get(): Record<string, unknown> },
    adapterName: string,
    sessionId?: string,
    sessionManager?: { getSession(id: string): { record?: { outputMode?: string } } | undefined },
  ): OutputMode {
    if (sessionId && sessionManager) {
      const session = sessionManager.getSession(sessionId);
      const mode = session?.record?.outputMode;
      if (mode === "low" || mode === "medium" || mode === "high") return mode as OutputMode;
    }
    const config = configManager.get();
    const channels = config.channels as Record<string, Record<string, unknown>> | undefined;
    const adapterConfig = channels?.[adapterName];
    const adapterMode = adapterConfig?.outputMode;
    if (adapterMode === "low" || adapterMode === "medium" || adapterMode === "high") return adapterMode as OutputMode;
    const globalMode = config.outputMode;
    if (globalMode === "low" || globalMode === "medium" || globalMode === "high") return globalMode as OutputMode;
    return "medium";
  }
}