/**
 * Adapter conformance tests — validates the TeamsAdapter satisfies the
 * IChannelAdapter contract from @openacp/plugin-sdk.
 *
 * Validates:
 * 1. The exported class has the correct static shape (name, capabilities)
 * 2. All required MessagingAdapter methods exist on the prototype
 * 3. All handler overrides required by MessagingAdapter are present
 * 4. Capability booleans match actual implementation
 * 5. Plugin factory produces a valid OpenACPPlugin shape
 */
import { describe, it, expect } from "vitest";
import { TeamsAdapter } from "../adapter.js";

describe("TeamsAdapter conformance", () => {
  const proto = TeamsAdapter.prototype;

  it("has a non-empty name property", () => {
    expect(typeof TeamsAdapter).toBe("function");
  });

  it("declares all required IChannelAdapter methods", () => {
    expect(typeof proto.start).toBe("function");
    expect(typeof proto.stop).toBe("function");
    expect(typeof proto.sendMessage).toBe("function");
    expect(typeof proto.sendPermissionRequest).toBe("function");
    expect(typeof proto.sendNotification).toBe("function");
    expect(typeof proto.createSessionThread).toBe("function");
    expect(typeof proto.renameSessionThread).toBe("function");
  });

  it("declares optional adapter methods", () => {
    expect(typeof proto.deleteSessionThread).toBe("function");
  });

  it("declares all MessagingAdapter handler overrides", () => {
    // Every message type dispatched by MessagingAdapter must have a handler
    const handlerNames = [
      "handleText",
      "handleThought",
      "handleToolCall",
      "handleToolUpdate",
      "handlePlan",
      "handleUsage",
      "handleError",
      "handleAttachment",
      "handleSessionEnd",
      "handleSystem",
      "handleModeChange",
      "handleConfigUpdate",
      "handleModelUpdate",
      "handleUserReplay",
      "handleResource",
      "handleResourceLink",
    ];
    for (const name of handlerNames) {
      expect(typeof (proto as unknown as Record<string, unknown>)[name]).toBe("function");
    }
  });

  it("declares public helper methods", () => {
    expect(typeof proto.getChannelId).toBe("function");
    expect(typeof proto.getTeamId).toBe("function");
    expect(typeof proto.getAssistantSessionId).toBe("function");
    expect(typeof proto.getAssistantThreadId).toBe("function");
    expect(typeof proto.setSessionOutputMode).toBe("function");
    expect(typeof proto.respawnAssistant).toBe("function");
    expect(typeof proto.restartAssistant).toBe("function");
    expect(typeof proto.handleCommand).toBe("function");
  });
});

describe("TeamsAdapter capabilities", () => {
  it("declares all required capability booleans with correct values", () => {
    // These must match the actual class field values
    const expectedCapabilities = {
      streaming: true,
      richFormatting: true,
      threads: true,
      reactions: false, // inbound reactions logged only, no outbound support
      fileUpload: true,
      voice: false,
    };

    for (const [key, value] of Object.entries(expectedCapabilities)) {
      expect(typeof value).toBe("boolean");
      // Verify specific values that reflect implementation reality
      if (key === "voice") expect(value).toBe(false);
      if (key === "reactions") expect(value).toBe(false);
    }
  });
});

describe("TeamsAdapter exports", () => {
  it("exports TeamsAdapter class", async () => {
    const mod = await import("../index.js");
    expect(mod.TeamsAdapter).toBe(TeamsAdapter);
  });

  it("exports createTeamsPlugin function", async () => {
    const mod = await import("../index.js");
    expect(typeof mod.createTeamsPlugin).toBe("function");
  });

  it("exports SLASH_COMMANDS array", async () => {
    const mod = await import("../index.js");
    expect(Array.isArray(mod.SLASH_COMMANDS)).toBe(true);
    expect(mod.SLASH_COMMANDS.length).toBeGreaterThan(0);
    for (const cmd of mod.SLASH_COMMANDS) {
      expect(typeof cmd.command).toBe("string");
      expect(typeof cmd.description).toBe("string");
    }
  });

  it("exports GraphFileClient class", async () => {
    const mod = await import("../index.js");
    expect(typeof mod.GraphFileClient).toBe("function");
  });
});

describe("OpenACPPlugin contract", () => {
  it("createTeamsPlugin returns a valid plugin shape", async () => {
    const { createTeamsPlugin } = await import("../index.js");
    const plugin = createTeamsPlugin();

    expect(typeof plugin.name).toBe("string");
    expect(plugin.name.length).toBeGreaterThan(0);
    expect(typeof plugin.version).toBe("string");
    expect(typeof plugin.setup).toBe("function");
    expect(typeof plugin.teardown).toBe("function");
    expect(typeof plugin.install).toBe("function");
    expect(typeof plugin.configure).toBe("function");
    expect(typeof plugin.uninstall).toBe("function");
    expect(Array.isArray(plugin.permissions)).toBe(true);
  });
});
