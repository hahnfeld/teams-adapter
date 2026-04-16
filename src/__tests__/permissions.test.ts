/**
 * Unit tests for PermissionHandler.
 */
import { describe, it, expect, vi, beforeEach } from "vitest";
import { PermissionHandler } from "../permissions.js";

function mockComposer() {
  const addPermission = vi.fn().mockReturnValue("entry-1");
  const resolvePermission = vi.fn();
  const msg = { addPermission, resolvePermission };
  return {
    getOrCreate: vi.fn().mockReturnValue(msg),
    get: vi.fn().mockReturnValue(msg),
    msg,
    addPermission,
    resolvePermission,
  };
}

describe("PermissionHandler", () => {
  let handler: PermissionHandler;
  let mockGetSession: ReturnType<typeof vi.fn>;
  let mockSendNotification: ReturnType<typeof vi.fn>;
  let composer: ReturnType<typeof mockComposer>;

  beforeEach(() => {
    mockGetSession = vi.fn();
    mockSendNotification = vi.fn().mockResolvedValue(undefined);
    composer = mockComposer();
    handler = new PermissionHandler(mockGetSession, mockSendNotification, composer as any);
  });

  describe("sendPermissionRequest", () => {
    it("adds a permission entry to the session card", async () => {
      const mockContext = {
        activity: { id: "a1", conversation: { id: "conv-1" } },
        sendActivity: vi.fn().mockResolvedValue({ id: "activity-1" }),
      };

      const session = {
        id: "session-1",
        name: "Test Session",
        permissionGate: { requestId: "req-1", resolve: vi.fn() },
      };

      const request = {
        id: "req-1",
        description: "Allow file read?",
        options: [
          { id: "allow", label: "Allow", isAllow: true },
          { id: "deny", label: "Deny", isAllow: false },
        ],
      };

      await handler.sendPermissionRequest(session as any, request, mockContext as any);

      // Permission entry added to the composer
      expect(composer.addPermission).toHaveBeenCalledWith(
        "Allow file read?",
        expect.arrayContaining([
          expect.objectContaining({ title: "✅ Allow", data: expect.objectContaining({ verb: "allow" }) }),
          expect.objectContaining({ title: "❌ Deny", data: expect.objectContaining({ verb: "deny" }) }),
        ]),
      );

      // Notification fired
      expect(mockSendNotification).toHaveBeenCalledWith(
        expect.objectContaining({ sessionId: "session-1", type: "permission" }),
      );
    });
  });

  describe("handleCardAction", () => {
    it("returns false for unknown verbs", async () => {
      const mockContext = { activity: { from: { id: "user-1" } } };
      const result = await handler.handleCardAction(mockContext as any, "unknown_verb", "session-1", "key", "req-1");
      expect(result).toBe(false);
    });

    it("returns true for unknown callback key (expired)", async () => {
      const mockContext = { activity: { from: { id: "user-1" } } };
      const result = await handler.handleCardAction(mockContext as any, "allow", "session-1", "nonexistent-key", "req-1");
      expect(result).toBe(true);
    });

    it("resolves permission and updates the entry in the card", async () => {
      const mockResolve = vi.fn();
      const session = {
        id: "session-1",
        permissionGate: { requestId: "req-1", resolve: mockResolve },
      };
      mockGetSession.mockReturnValue(session);

      // Send a permission request to create a pending entry
      const sendContext = {
        activity: { id: "a1", conversation: { id: "conv-1" } },
        sendActivity: vi.fn().mockResolvedValue({ id: "activity-1" }),
      };
      await handler.sendPermissionRequest(session as any, {
        id: "req-1",
        description: "Allow?",
        options: [
          { id: "opt-allow", label: "Allow", isAllow: true },
          { id: "opt-deny", label: "Deny", isAllow: false },
        ],
      }, sendContext as any);

      // Extract the callbackKey from the addPermission call
      const actions = composer.addPermission.mock.calls[0][1];
      const callbackKey = actions[0].data.callbackKey;

      // Handle the card action
      const actionContext = { activity: { from: { id: "user-1", name: "Test User" } } };
      const result = await handler.handleCardAction(actionContext as any, "allow", "session-1", callbackKey, "req-1");

      expect(result).toBe(true);
      expect(mockResolve).toHaveBeenCalledWith("opt-allow");
      expect(composer.resolvePermission).toHaveBeenCalledWith(
        "entry-1",
        expect.stringContaining("Allowed"),
      );
    });
  });

  describe("evictStale", () => {
    it("does not crash with many pending entries", async () => {
      const mockContext = {
        activity: { id: "a1", conversation: { id: "c1" } },
        sendActivity: vi.fn().mockResolvedValue({ id: "a1" }),
      };
      const session = { id: "s1", name: "Test" };
      const request = {
        id: "r1",
        description: "Allow?",
        options: [{ id: "o1", label: "Allow", isAllow: true }],
      };

      for (let i = 0; i < 110; i++) {
        await handler.sendPermissionRequest(session as any, request, mockContext as any);
      }

      expect(composer.addPermission).toHaveBeenCalledTimes(110);
    });
  });
});
