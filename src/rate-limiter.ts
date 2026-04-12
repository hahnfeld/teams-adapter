/**
 * Per-conversation rate limiter for Teams Bot Framework API.
 *
 * Teams enforces rate limits per bot per conversation thread:
 *   7 ops/1s, 8 ops/2s, 60 ops/30s, 1800 ops/3600s
 *
 * This limiter tracks sliding windows and queues operations when any
 * window would be exceeded. Supports coalescing — queued operations
 * with the same key replace earlier ones (e.g., rapid message edits
 * collapse into a single PUT).
 */
import { log } from "@openacp/plugin-sdk";

/** Sliding window definition: max operations allowed within a time window. */
interface WindowConfig {
  windowMs: number;
  max: number;
}

/** Our targets — slightly under Teams limits to leave headroom. */
const RATE_WINDOWS: WindowConfig[] = [
  { windowMs: 1_000, max: 6 },
  { windowMs: 2_000, max: 7 },
  { windowMs: 30_000, max: 55 },
  { windowMs: 3_600_000, max: 1_700 },
];

interface QueuedOp<T = unknown> {
  fn: () => Promise<T>;
  key?: string;
  resolve: (value: T | undefined) => void;
  reject: (err: unknown) => void;
}

/** Per-conversation queue state. */
class ConversationQueue {
  /** Timestamps of completed operations (for sliding window tracking). */
  private timestamps: number[] = [];
  private queue: QueuedOp[] = [];
  private draining = false;
  private destroyed = false;
  private pausedUntil = 0;
  private drainTimer?: ReturnType<typeof setTimeout>;
  /** Resolves the inner drain-wait promise on destroy, preventing coroutine leak. */
  private drainWaitResolve?: () => void;

  /** Enqueue an operation. If key matches a pending op, replace it (coalescing). */
  enqueue<T>(fn: () => Promise<T>, key?: string): Promise<T | undefined> {
    if (this.destroyed) return Promise.resolve(undefined);
    return new Promise<T | undefined>((resolve, reject) => {
      if (key) {
        const idx = this.queue.findIndex((op) => op.key === key);
        if (idx !== -1) {
          // Coalesce: resolve the old op as undefined (skipped), replace with new
          this.queue[idx].resolve(undefined);
          this.queue[idx] = { fn, key, resolve: resolve as (v: unknown) => void, reject };
          return;
        }
      }
      this.queue.push({ fn, key, resolve: resolve as (v: unknown) => void, reject });
      this.scheduleDrain();
    });
  }

  get pending(): number {
    return this.queue.length;
  }

  destroy(): void {
    this.destroyed = true;
    if (this.drainTimer) { clearTimeout(this.drainTimer); this.drainTimer = undefined; }
    // Unblock any suspended drain-wait so the coroutine can exit
    if (this.drainWaitResolve) { this.drainWaitResolve(); this.drainWaitResolve = undefined; }
    // Resolve all pending as undefined
    for (const op of this.queue) op.resolve(undefined);
    this.queue.length = 0;
  }

  private scheduleDrain(): void {
    if (this.draining) return;
    if (this.drainTimer) return;

    const delay = this.getDelay();
    this.drainTimer = setTimeout(() => {
      this.drainTimer = undefined;
      this.drain();
    }, delay);
  }

  /** Calculate how long to wait before next op is allowed. */
  private getDelay(): number {
    const now = Date.now();

    // Respect 429 pause
    if (now < this.pausedUntil) {
      return this.pausedUntil - now;
    }

    let maxDelay = 0;
    for (const { windowMs, max } of RATE_WINDOWS) {
      const cutoff = now - windowMs;
      // Timestamps are in ascending order — find the first one in the window
      const firstIdx = this.timestamps.findIndex((t) => t > cutoff);
      if (firstIdx === -1) continue;
      const opsInWindow = this.timestamps.length - firstIdx;
      if (opsInWindow >= max) {
        const oldest = this.timestamps[firstIdx];
        const wait = oldest + windowMs - now + 1;
        maxDelay = Math.max(maxDelay, wait);
      }
    }
    return maxDelay;
  }

  private async drain(): Promise<void> {
    if (this.draining) return;
    this.draining = true;

    try {
      while (this.queue.length > 0 && !this.destroyed) {
        const delay = this.getDelay();
        if (delay > 0) {
          await new Promise<void>((r) => {
            this.drainWaitResolve = r;
            this.drainTimer = setTimeout(() => {
              this.drainTimer = undefined;
              this.drainWaitResolve = undefined;
              r();
            }, delay);
          });
          // Check if destroyed while waiting
          if (this.destroyed) break;
          continue;
        }

        const op = this.queue.shift()!;

        try {
          const result = await op.fn();
          // Record timestamp only on success — 429s should not consume quota
          const now = Date.now();
          this.timestamps.push(now);
          this.pruneTimestamps(now);
          op.resolve(result);
        } catch (err: unknown) {
          const statusCode = (err as { statusCode?: number })?.statusCode;
          if (statusCode === 429) {
            // Do NOT record a timestamp — the call was rejected
            const retryAfterRaw = (err as { headers?: Record<string, string> })?.headers?.["retry-after"];
            const retryAfterSec = retryAfterRaw ? parseInt(retryAfterRaw, 10) : NaN;
            const retryMs = !isNaN(retryAfterSec) && retryAfterSec > 0
              ? retryAfterSec * 1000
              : 2000;
            log.warn({ retryMs }, "[RateLimiter] 429 received, pausing");
            this.pausedUntil = Date.now() + retryMs;
            // Re-queue the failed op at the front
            this.queue.unshift(op);
            continue;
          }
          // Non-429 errors still consumed a slot
          const now = Date.now();
          this.timestamps.push(now);
          this.pruneTimestamps(now);
          op.reject(err);
        }
      }
    } finally {
      this.draining = false;
    }
  }

  /** Remove timestamps older than the largest window. */
  private pruneTimestamps(now: number): void {
    const maxWindow = RATE_WINDOWS[RATE_WINDOWS.length - 1].windowMs;
    const cutoff = now - maxWindow;
    while (this.timestamps.length > 0 && this.timestamps[0] <= cutoff) {
      this.timestamps.shift();
    }
  }
}

/**
 * Manages per-conversation rate-limited queues.
 *
 * Usage:
 *   const limiter = new ConversationRateLimiter();
 *   await limiter.enqueue(conversationId, () => sendText(ctx, text), "main-msg");
 */
export class ConversationRateLimiter {
  private queues = new Map<string, ConversationQueue>();

  /** Enqueue an operation for a conversation. Key enables coalescing. */
  enqueue<T>(conversationId: string, fn: () => Promise<T>, key?: string): Promise<T | undefined> {
    let queue = this.queues.get(conversationId);
    if (!queue) {
      queue = new ConversationQueue();
      this.queues.set(conversationId, queue);
    }
    return queue.enqueue(fn, key);
  }

  /** Clean up a conversation's queue (e.g., on session end). */
  cleanup(conversationId: string): void {
    const queue = this.queues.get(conversationId);
    if (queue) {
      queue.destroy();
      this.queues.delete(conversationId);
    }
  }

  /** Clean up all queues. */
  destroy(): void {
    for (const queue of this.queues.values()) queue.destroy();
    this.queues.clear();
  }
}
