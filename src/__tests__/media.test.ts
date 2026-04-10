/**
 * Unit tests for media utilities.
 */
import { describe, it, expect } from "vitest";
import { isAttachmentTooLarge, buildFileAttachmentCard } from "../media.js";

describe("isAttachmentTooLarge", () => {
  it("returns false for small files", () => {
    expect(isAttachmentTooLarge(1024)).toBe(false);
  });

  it("returns false at exactly the limit", () => {
    expect(isAttachmentTooLarge(250 * 1024 * 1024)).toBe(false);
  });

  it("returns true above the limit", () => {
    expect(isAttachmentTooLarge(250 * 1024 * 1024 + 1)).toBe(true);
  });
});

describe("buildFileAttachmentCard", () => {
  it("returns an Adaptive Card v1.2", () => {
    const card = buildFileAttachmentCard("test.txt", 1024, "text/plain");
    expect(card.type).toBe("AdaptiveCard");
    expect(card.version).toBe("1.2");
  });

  it("includes file name in card body", () => {
    const card = buildFileAttachmentCard("report.pdf", 2048, "application/pdf");
    const textBlocks = card.body.filter((b: any) => b.type === "TextBlock");
    const texts = textBlocks.map((b: any) => b.text).join(" ");
    expect(texts).toContain("report.pdf");
  });

  it("formats size in KB for small files", () => {
    const card = buildFileAttachmentCard("small.txt", 512, "text/plain");
    const allText = JSON.stringify(card.body);
    expect(allText).toContain("KB");
  });

  it("formats size in MB for large files", () => {
    const card = buildFileAttachmentCard("big.zip", 5 * 1024 * 1024, "application/zip");
    const allText = JSON.stringify(card.body);
    expect(allText).toContain("MB");
  });

  it("does not include file:// URLs", () => {
    const card = buildFileAttachmentCard("test.txt", 1024, "text/plain");
    const serialized = JSON.stringify(card);
    expect(serialized).not.toContain("file://");
  });
});
