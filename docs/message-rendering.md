# Message Rendering Reference

How each OpenACP message type is rendered in the Teams Adaptive Card.

---

## Card Structure

Every session has a single Adaptive Card that is created on the first event and updated in-place via REST API PUT. All message types render into this card — nothing is sent as a standalone message.

The card is flushed on a 500ms debounce. Updates are rate-limited per conversation.

---

## Entry Types

### Title

Bold session name. Always the first entry. Set once via auto-naming, updated if the session is renamed.

```
│ Fix rate limiter bug                             │
```

Adaptive Card: `TextBlock`, weight Bolder, size Medium, fontType Monospace.

---

### Timed (Tool, Thinking)

Two-level Container with a live elapsed timer. Used for `tool_call` and `thought` events.

**While running** — level 1 only, timer ticks every 1 second:

```
│ 🔧 Read src/main.ts…  (1.2s)                    │
│ ☁️ Thinking…  (0.8s)                             │
```

**After completion** — level 1 + level 2 result via 3-column ColumnSet:

```
│ 🔧 Read src/main.ts                             │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ Read src/main.ts (216 lines)       │    │
│ │  │   │ (0.3s)                             │    │
│ └──┴───┴────────────────────────────────────┘    │
│                                                  │
│ ☁️ Thinking                                      │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ Need to check the sliding window   │    │
│ │  │   │ implementation  (2.1s)             │    │
│ └──┴───┴────────────────────────────────────┘    │
```

**Tool lifecycle:**
1. `tool_call` event → `addTimedStart("🔧", summary)` — creates entry, starts timer
2. `tool_update` event → `addTimedResult(id, summary)` — sets result, stops timer

**Thinking lifecycle:**
1. `thought` event → `addThinking(text)` — creates entry on first call, accumulates text on subsequent calls
2. Any non-thought event (`text`, `tool_call`, `session_end`, etc.) → `closeActiveThinking()` — sets accumulated text as result, stops timer

Level 2 uses a 3-column ColumnSet for true indentation — wrapped text stays indented:
- Column 1: 20px spacer
- Column 2: `⎿` (auto width, top-aligned)
- Column 3: content text (stretch, wrap: true)

---

### Info (Error, System, Mode, Config, Model)

Two-level Container, no timer. One-shot entries for status changes and errors.

```
│ ❌ Error                                         │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ Failed to read file: permission    │    │
│ │  │   │ denied for /etc/shadow             │    │
│ └──┴───┴────────────────────────────────────┘    │
│                                                  │
│ ⚙️ Mode                                          │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ architect                          │    │
│ └──┴───┴────────────────────────────────────┘    │
│                                                  │
│ ⚙️ Model                                         │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ claude-sonnet-4-5                  │    │
│ └──┴───┴────────────────────────────────────┘    │
│                                                  │
│ ⚙️ Config                                        │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ outputMode                         │    │
│ └──┴───┴────────────────────────────────────┘    │
│                                                  │
│ ⚙️ System                                        │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ Session timeout extended to 120    │    │
│ │  │   │ minutes                            │    │
│ └──┴───┴────────────────────────────────────┘    │
```

Same 3-column ColumnSet as timed entries for the level 2 line.

| Message type   | Emoji | Label    | Detail source                    |
|----------------|-------|----------|----------------------------------|
| `error`        | ❌    | Error    | `content.text`                   |
| `system_message` | ⚙️  | System   | `content.text`                   |
| `mode_change`  | ⚙️    | Mode     | `metadata.modeId`                |
| `config_update`| ⚙️    | Config   | `metadata.key` or "updated"      |
| `model_update` | ⚙️    | Model    | `metadata.modelId`               |

---

### Text

Root-level streamed agent response. Always appended at the bottom, never nested under a tool or thinking entry.

```
│ I fixed the sliding window on line 42. The       │
│ burst counter now resets correctly after the      │
│ window expires.                                  │
```

Adaptive Card: `TextBlock`, size Small, fontType Monospace, wrap: true.

Consecutive text chunks are concatenated into the same entry. Auto-splits into a new card at 25,000 characters.

---

### Plan

Formatted checklist, updated in place on each `plan` event. No Container wrapper.

```
│ 📋 Plan                                         │
│ ✅ 1. Set up project structure                   │
│ ✅ 2. Create database schema                     │
│ 🔄 3. Implement API endpoints                    │
│ ⏳ 4. Write tests                                │
│ ⏳ 5. Deploy to staging                           │
```

Status icons: ✅ completed, 🔄 in_progress, ⏳ pending.

Singleton — `setPlan()` replaces the existing plan entry if one exists.

---

### Resource

Inline `📎` line for attachments, resources, and resource links.

```
│ 📎 report.pdf (24KB)                             │
│ 📎 [report.pdf](https://sharepoint.com/...)      │
│ 📎 API Documentation                             │
│ 📎 [API Docs](https://docs.example.com)          │
│ 📎 ⚠️ File too large (15MB): database.sql        │
```

Adaptive Card: `TextBlock`, size Small, fontType Monospace, wrap: true.

| Message type    | Format                                        |
|-----------------|-----------------------------------------------|
| `attachment`    | `📎 [fileName](shareUrl)` or `📎 fileName (size)` |
| `resource`      | `📎 {content.text}`                           |
| `resource_link` | `📎 [name](url)` or `📎 {content.text}`       |
| `user_replay`   | Plain text via `addText()` (no 📎 prefix)     |

---

### Usage

Italic footer. Singleton — replaced on each `usage` event. Always the last entry.

```
│ 13k tokens · 4.2s · $0.0312 · Task completed    │
```

Adaptive Card: `TextBlock`, isSubtle: true, size Small, fontType Monospace.

On `session_end`, "Task completed" is appended to the existing usage text.

---

### Divider

Horizontal rule. Used by the stall timer when a response is cut short.

```
│ ──────────────────────────────                   │
```

---

## Outside the Card

### Permission Requests

Sent as a separate Adaptive Card via `sendPermissionRequest()`. Uses the same Container + ColumnSet style.

**Pending:**
```
┌──────────────────────────────────────────────────┐
│ 🔐 Permission                                   │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ Allow file write to src/main.ts?   │    │
│ └──┴───┴────────────────────────────────────┘    │
│  [✅ Allow]  [❌ Deny]                           │
└──────────────────────────────────────────────────┘
```

Buttons are inside an `ActionSet` in the Container body (compact inline style).

**Responded (updated in-place):**
```
┌──────────────────────────────────────────────────┐
│ ✅ Permission — Allowed                          │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ Allow file write to src/main.ts?   │    │
│ │  │   │ (Matt, 3s)                         │    │
│ └──┴───┴────────────────────────────────────┘    │
└──────────────────────────────────────────────────┘
```

### Notifications

Sent to the notification channel as a separate Adaptive Card. Same info Container style.

```
┌──────────────────────────────────────────────────┐
│ ✅ Completed                                     │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ Fix rate limiter — Task finished   │    │
│ │  │   │ Open →                             │    │
│ └──┴───┴────────────────────────────────────┘    │
└──────────────────────────────────────────────────┘
```

| Type             | Emoji | Label            |
|------------------|-------|------------------|
| `completed`      | ✅    | Completed        |
| `error`          | ❌    | Error            |
| `permission`     | 🔐    | Permission       |
| `input_required` | 💬    | Input Required   |
| `budget_warning` | ⚠️    | Budget Warning   |

---

## Full Card Example

A realistic session showing multiple entry types in order:

```
┌──────────────────────────────────────────────────┐
│ Fix rate limiter bug                             │
│                                                  │
│ ☁️ Thinking                                      │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ Need to check the sliding window   │    │
│ │  │   │ implementation  (2.1s)             │    │
│ └──┴───┴────────────────────────────────────┘    │
│                                                  │
│ 📋 Plan                                         │
│ ✅ 1. Review rate limiter code                   │
│ ✅ 2. Fix sliding window bug                    │
│ 🔄 3. Run tests                                 │
│ ⏳ 4. Update docs                                │
│                                                  │
│ 🔧 Read src/rate-limiter.ts                      │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ Read src/rate-limiter.ts (216      │    │
│ │  │   │ lines)  (0.3s)                     │    │
│ └──┴───┴────────────────────────────────────┘    │
│                                                  │
│ 🔧 Edit src/rate-limiter.ts                      │
│ ┌──┬───┬────────────────────────────────────┐    │
│ │  │ ⎿ │ Edit src/rate-limiter.ts  (1.1s)   │    │
│ └──┴───┴────────────────────────────────────┘    │
│                                                  │
│ 🔧 Run: pnpm test…  (4.2s)                      │
│                                                  │
│ I fixed the sliding window on line 42. The       │
│ burst counter now resets correctly after the      │
│ window expires.                                  │
│                                                  │
│ 📎 [test-results.txt](https://sharepoint/...)    │
│                                                  │
│ 13k tokens · 8.1s · $0.0512 · Task completed    │
└──────────────────────────────────────────────────┘
```
