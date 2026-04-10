/**
 * Unit tests for formatting utilities.
 */
import { describe, it, expect } from "vitest";
import {
  progressBar,
  formatTokens,
  truncateContent,
  splitMessage,
  extractContentText,
  stripCodeFences,
  resolveToolIcon,
  formatToolCall,
  formatPlan,
  formatUsage,
  renderUsageCard,
} from "../formatting.js";

describe("progressBar", () => {
  it("renders empty bar for 0", () => {
    expect(progressBar(0)).toBe("░░░░░░░░░░");
  });

  it("renders full bar for 1", () => {
    expect(progressBar(1)).toBe("▓▓▓▓▓▓▓▓▓▓");
  });

  it("renders half bar for 0.5", () => {
    expect(progressBar(0.5)).toBe("▓▓▓▓▓░░░░░");
  });

  it("clamps values above 1", () => {
    expect(progressBar(2)).toBe("▓▓▓▓▓▓▓▓▓▓");
  });

  it("clamps values below 0", () => {
    expect(progressBar(-0.5)).toBe("░░░░░░░░░░");
  });

  it("respects custom length", () => {
    expect(progressBar(0.5, 4)).toBe("▓▓░░");
  });
});

describe("formatTokens", () => {
  it("formats small numbers as-is", () => {
    expect(formatTokens(500)).toBe("500");
  });

  it("formats thousands with k suffix", () => {
    expect(formatTokens(1000)).toBe("1k");
    expect(formatTokens(1500)).toBe("2k");
    expect(formatTokens(200000)).toBe("200k");
  });
});

describe("truncateContent", () => {
  it("returns short text unchanged", () => {
    expect(truncateContent("hello", 100)).toBe("hello");
  });

  it("truncates long text with suffix", () => {
    const result = truncateContent("a".repeat(100), 50);
    expect(result.length).toBe(50);
    expect(result.endsWith("… (truncated)")).toBe(true);
  });
});

describe("splitMessage", () => {
  it("returns single chunk for short text", () => {
    expect(splitMessage("hello", 100)).toEqual(["hello"]);
  });

  it("splits on paragraph boundaries", () => {
    const text = "paragraph 1\n\nparagraph 2\n\nparagraph 3";
    const chunks = splitMessage(text, 20);
    expect(chunks.length).toBeGreaterThan(1);
    // All chunks should be within limit
    for (const chunk of chunks) {
      expect(chunk.length).toBeLessThanOrEqual(20);
    }
  });

  it("handles text with no paragraph breaks", () => {
    const text = "a".repeat(50);
    const chunks = splitMessage(text, 20);
    expect(chunks.length).toBeGreaterThanOrEqual(1);
  });
});

describe("extractContentText", () => {
  it("returns empty string for null/undefined", () => {
    expect(extractContentText(null)).toBe("");
    expect(extractContentText(undefined)).toBe("");
  });

  it("returns string directly", () => {
    expect(extractContentText("hello")).toBe("hello");
  });

  it("extracts text from array of objects", () => {
    expect(extractContentText([{ text: "a" }, { text: "b" }])).toBe("a\nb");
  });

  it("extracts text from object with text property", () => {
    expect(extractContentText({ text: "hello" })).toBe("hello");
  });
});

describe("stripCodeFences", () => {
  it("removes code fences", () => {
    expect(stripCodeFences("```js\nconsole.log('hi')\n```")).toBe("console.log('hi')");
  });

  it("leaves text without fences unchanged", () => {
    expect(stripCodeFences("no fences here")).toBe("no fences here");
  });
});

describe("resolveToolIcon", () => {
  it("uses status icon when available", () => {
    expect(resolveToolIcon("read", undefined, "completed")).toBe("✅");
  });

  it("uses kind icon", () => {
    expect(resolveToolIcon("read")).toBe("📖");
    expect(resolveToolIcon("execute")).toBe("▶️");
    expect(resolveToolIcon("search")).toBe("🔍");
  });

  it("falls back to wrench", () => {
    expect(resolveToolIcon("unknown_kind")).toBe("🔧");
  });
});

describe("formatToolCall", () => {
  it("renders minimal tool call", () => {
    const result = formatToolCall({ id: "t1", name: "Read" });
    expect(result).toContain("**Read**");
  });

  it("includes viewer links when present", () => {
    const result = formatToolCall({
      id: "t1",
      name: "Edit",
      viewerLinks: { file: "http://example.com/file" },
      viewerFilePath: "src/index.ts",
    });
    expect(result).toContain("[View index.ts]");
  });

  it("shows input/output in high verbosity", () => {
    const result = formatToolCall(
      { id: "t1", name: "Read", rawInput: { file_path: "/foo" }, content: "file content" },
      "high",
    );
    expect(result).toContain("**Input:**");
    expect(result).toContain("**Output:**");
  });
});

describe("formatPlan", () => {
  it("renders summary in medium verbosity", () => {
    const entries = [
      { content: "Step 1", status: "completed" },
      { content: "Step 2", status: "pending" },
    ];
    const result = formatPlan(entries, "medium");
    expect(result).toContain("1/2 steps completed");
  });

  it("renders full plan in high verbosity", () => {
    const entries = [
      { content: "Step 1", status: "completed" },
      { content: "Step 2", status: "in_progress" },
    ];
    const result = formatPlan(entries, "high");
    expect(result).toContain("✅");
    expect(result).toContain("🔄");
    expect(result).toContain("Step 1");
    expect(result).toContain("Step 2");
  });
});

describe("formatUsage", () => {
  it("returns unavailable when no tokens", () => {
    expect(formatUsage({})).toContain("unavailable");
  });

  it("formats token count in medium", () => {
    expect(formatUsage({ tokensUsed: 5000 }, "medium")).toContain("5k tokens");
  });

  it("includes cost when available", () => {
    expect(formatUsage({ tokensUsed: 1000, cost: 0.05 }, "medium")).toContain("$0.05");
  });

  it("shows progress bar in high verbosity", () => {
    const result = formatUsage({ tokensUsed: 50000, contextSize: 200000 }, "high");
    expect(result).toContain("▓");
    expect(result).toContain("%");
  });
});

describe("renderUsageCard", () => {
  it("returns card body with TextBlock", () => {
    const { body } = renderUsageCard({ tokensUsed: 1000 }, "medium");
    expect(body).toHaveLength(1);
    expect((body[0] as { type: string }).type).toBe("TextBlock");
  });
});
