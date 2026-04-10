# Teams Adapter Code Review Report

**Reviewer**: Principal Engineer Review
**Date**: 2026-04-10
**Scope**: Full adapter review against OpenACP adapter spec, Telegram reference implementation, and MS Teams production best practices

---

## Remediation Summary

All P0 blockers and most P1/P2 issues have been fixed. Key changes:

- **Per-session dispatch queues** matching Telegram's `_dispatchQueues` pattern for serialized event delivery
- **Command routing wired** — slash commands now detected and dispatched before `core.handleMessage()`
- **Rate limiting with retry** — `sendActivityWithRetry()` with exponential backoff + jitter for 429/502/504
- **Idempotency handling** — duplicate activity ID detection for Teams 15-second retries
- **Graph API file support** — `GraphFileClient` for authenticated downloads and OneDrive uploads with sharing links
- **Plugin entry fixed** — `setup()` now calls `adapter.start()`, `index.ts` moved into `src/`
- **Missing dependencies added** — `@microsoft/teams.api`, `nanoid`
- **Capabilities corrected** — `voice: false`, `maxMessageLength: 28000`
- **Assistant lifecycle** — `setupAssistant()` called in `start()` when configured
- **Adaptive Cards v1.2** baseline for mobile compatibility
- **Test suite added** — conformance, formatting, media, permissions tests
- **Duplicate types cleaned** — `MessageRef` deduped, redundant `appendText` removed

### Original Severity Breakdown (pre-fix)
- **P0 (Blockers)**: 7 issues -> **0 remaining**
- **P1 (Critical)**: 8 issues -> **2 remaining** (real thread creation, proactive messaging)
- **P2 (Important)**: 6 issues -> **3 remaining** (stub commands, some `as any` casts)
- **P3 (Minor)**: 4 issues -> **1 remaining** (`adaptivecards-templating` unused)

---

## P0 — Production Blockers

### 1. No Test Coverage
**Files**: None exist
**Impact**: Cannot verify adapter conformance, cannot prevent regressions
**OpenACP Requirement**: All adapters must pass `runAdapterConformanceTests()` from `@openacp/plugin-sdk/testing`. The Telegram adapter has 10+ test files covering conformance, startup, formatting, control messages, and prerequisites.
**Fix**: Add conformance tests, unit tests for formatting/permissions/draft-manager, and integration tests for message routing.

### 2. `index.ts` Outside `rootDir`
**File**: `index.ts` (root) + `tsconfig.json`
**Impact**: Build will fail — TypeScript won't compile files outside `rootDir: "./src"`
**Fix**: Move plugin entry into `src/` or adjust `rootDir`/`include` in tsconfig.

### 3. Plugin Entry Doesn't Call `adapter.start()`
**File**: `index.ts:18-20`
**Impact**: The adapter's message handlers and Teams App are never initialized. The bot won't receive any messages.
**Reference**: Telegram plugin calls `adapter.start()` in its `setup()` hook.
**Fix**: Call `await adapter.start()` after construction in plugin `setup()`.

### 4. Missing Dependencies in `package.json`
**File**: `package.json`
**Missing**:
- `@microsoft/teams.api` — imported in `adapter.ts:5` for `InvokeResponse`
- `nanoid` — imported in `permissions.ts:3`
**Fix**: Add both to `dependencies`.

### 5. Command Routing Dead — Messages Never Reach Slash Commands
**File**: `adapter.ts:153-238`
**Impact**: The message handler processes all text as regular messages. It never checks for the `/` prefix or calls `handleCommand()`. All slash commands (`/new`, `/cancel`, `/status`, etc.) are unreachable.
**Fix**: Add command detection before `core.handleMessage()`.

### 6. File Downloads Don't Authenticate
**File**: `media.ts:12-13`
**Impact**: `fetch(url)` for Teams file attachments will fail with 401. Teams attachment URLs require the bot's Bearer token.
**Reference**: Teams Bot Framework requires authorization headers for content URLs.
**Fix**: Pass bot token or use the TurnContext's connector client to download files.

### 7. File Attachment Cards Use `file://` Server Paths
**File**: `media.ts` via `adapter.ts:674`
**Impact**: `file://${attachment.filePath}` points to a server-local path. Teams users cannot access server filesystem paths.
**Fix**: Upload files to SharePoint/OneDrive via Graph API, or serve via the adapter's HTTP endpoint and provide an accessible URL.

---

## P1 — Critical Issues

### 8. `sendMessage` Context Lifecycle Race Condition
**File**: `adapter.ts:487-509`
**Impact**: Context is stored in `_sessionContexts`, then `super.sendMessage()` dispatches to handler methods that read it. The `finally` block deletes the context with only a `setTimeout(0)` guard. Any handler awaiting real I/O (network calls to Teams) could lose its context mid-execution.
**Fix**: Use a reference-counting approach or don't delete context in `sendMessage()` — let the message handler's scope manage it.

### 9. `createSessionThread` Returns Fake Thread IDs
**File**: `adapter.ts:748-761`
**Impact**: Returns `thread-${sessionId}-${Date.now()}` instead of creating a real Teams conversation. Session-to-thread mapping is synthetic; messages can't actually be routed to separate threads.
**Fix**: Implement actual Teams channel thread creation via Bot Framework `createConversation` API, or document that Teams adapter operates in single-conversation mode and adjust `capabilities.threads`.

### 10. No Rate Limiting or Retry Logic
**Impact**: Teams enforces strict rate limits (7 msg/sec per conversation, 50 RPS global per tenant). No 429 handling, no exponential backoff, no Retry-After header parsing.
**Reference**: Microsoft docs mandate retry with exponential backoff + jitter for 429, 412, 502, 504.
**Fix**: Add retry middleware to `SendQueue` with Teams-specific limits.

### 11. No Idempotency Handling
**Impact**: Teams retries requests if bot takes >15 seconds to respond. Without deduplication, the adapter processes the same message twice.
**Fix**: Track processed activity IDs and skip duplicates.

### 12. `setupAssistant()` Never Called
**File**: `adapter.ts:403-426`
**Impact**: The assistant feature is dead code. `start()` doesn't call `setupAssistant()`, so `this.assistantSession` is always null.
**Fix**: Call `setupAssistant()` in `start()` when `assistantThreadId` is configured, or document it as opt-in.

### 13. `voice: true` Capability With No Voice Implementation
**File**: `adapter.ts:41`
**Impact**: Adapter declares voice capability but has no voice handling. Core may route voice content expecting it to work.
**Fix**: Set `voice: false` until voice is implemented, or implement Teams voice support.

### 14. Proactive Messaging Not Implemented
**Impact**: `sendNotification` only works if `notificationContext` was captured from a previous inbound message. If the bot restarts or no message has been received in the notification channel, notifications are silently dropped.
**Fix**: Store conversation references persistently and use `adapter.continueConversation()` for proactive messaging.

### 15. Adaptive Card Version — Mobile Compatibility
**File**: Multiple — cards use version "1.4"
**Impact**: Teams mobile only reliably supports Adaptive Card schema v1.2. Cards using 1.4 features may not render on mobile.
**Fix**: Use v1.2 as the baseline, or feature-detect and downgrade for mobile contexts.

---

## P2 — Important Issues

### 16. Duplicate `MessageRef` Interface
**Files**: `activity.ts:22-25` and `draft-manager.ts:8-11`
**Fix**: Define once and import.

### 17. Redundant Text Buffering in `handleText`
**File**: `adapter.ts:529-531`
**Impact**: `draft.append(content.text)` and `draftManager.appendText(sessionId, content.text)` both buffer text, creating divergent state.
**Fix**: Remove the redundant `appendText` call or unify the buffering.

### 18. `sendNotification` Silently Drops When No Context
**File**: `adapter.ts:721-744`
**Impact**: If no message has been received in the notification channel, all notifications are silently logged at debug level and discarded.
**Fix**: Queue notifications for delivery when context becomes available, or use proactive messaging.

### 19. 47+ `as any` Type Casts
**Files**: Throughout codebase
**Impact**: Defeats TypeScript's type safety. Likely caused by SDK type mismatches.
**Fix**: Create proper type definitions or use type-safe wrapper functions. Investigate SDK version compatibility.

### 20. `handleToolUpdate` Is a No-Op
**File**: `adapter.ts:552-556`
**Impact**: Tool progress updates (status changes, partial output) are silently discarded. Users get no visibility into tool execution progress.
**Fix**: Implement tool update rendering, at minimum for high verbosity mode.

### 21. Many Commands Are Stub Implementations
**Files**: `commands/admin.ts`, `commands/agents.ts`, `commands/doctor.ts`, `commands/integrate.ts`, `commands/menu.ts`, `commands/settings.ts`, `commands/new-session.ts`
**Impact**: 15+ commands respond with "not yet implemented". Users will encounter dead functionality.
**Fix**: Implement core commands (`/new`, `/sessions`, `/agents`, `/menu`, `/settings`, `/doctor`) or remove stubs.

---

## P3 — Minor Issues

### 22. `handleText` — Catches Error But Logs and Returns
Correct pattern, but `getSessionContext()` throwing is expected when context is missing. Consider using `getContext()` (which returns null) instead of try/catch on `getSessionContext()`.

### 23. Dynamic Import in Hot Path
**File**: `adapter.ts:569, 595`
`await import("./formatting.js")` in `handlePlan` and `handleUsage` creates repeated dynamic imports in message handlers. These should be static imports.

### 24. `maxMessageLength: 2000` — Too Low
**File**: `adapter.ts:64`
Teams supports messages up to ~28K characters (or ~100KB). 2000 is unnecessarily restrictive and will cause excessive message splitting.

### 25. `adaptivecards-templating` Dependency Unused
**File**: `package.json:28`
The package is listed as a dependency but never imported. Either use it for card templating (recommended) or remove it.

---

## Comparison with Telegram Reference Implementation

| Feature | Telegram | Teams | Gap |
|---------|----------|-------|-----|
| Conformance tests | Yes | **No** | Critical |
| Command routing | Bot commands + text prefix | **Dead** — not wired | Critical |
| Thread management | Forum topics via API | Fake IDs | Critical |
| Permission round-trip | Callback queries → resolve | Card actions → resolve | Working |
| Draft/streaming | Edit messages in-place | Update activities | Working |
| Rate limiting | grammY built-in | **None** | Critical |
| Proactive messaging | sendMessage to chatId | **Not implemented** | Significant |
| File handling | Download + re-upload | **Broken** (no auth, file:// URLs) | Critical |
| Voice/TTS | Edge TTS plugin | **Stub only** | Moderate |
| Error recovery | Resilient start even on prereq failure | Throws on init failure | Moderate |

---

## Recommendations (Priority Order)

1. Fix P0 blockers (build, deps, command routing, file handling)
2. Add conformance tests + unit tests
3. Implement rate limiting and retry logic
4. Implement proactive messaging for notifications
5. Fix context lifecycle race condition
6. Implement real thread creation or set `threads: false`
7. Set `voice: false` until implemented
8. Implement core slash commands (`/new`, `/sessions`, `/agents`)
9. Reduce `as any` usage with proper type definitions
10. Add idempotency handling for Teams retries
