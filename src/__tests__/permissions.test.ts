/**
 * Unit tests for PermissionHandler.
 */
import { describe, it, expect, vi, beforeEach } from "vitest";
import { PermissionHandler } from "../permissions.js";

describe("PermissionHandler", () => {
  let handler: PermissionHandler;
  let mockGetSession: ReturnType<typeof vi.fn>;
  let mockSendNotification: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    mockGetSession = vi.fn();
    mockSendNotification = vi.fn().mockResolvedValue(undefined);
    handler = new PermissionHandler(mockGetSession, mockSendNotification);
  });

  describe("sendPermissionRequest", () => {
    it("sends an Adaptive Card with permission buttons", async () => {
      const mockSendActivity = vi.fn().mockResolvedValue({ id: "activity-1" });
      const mockContext = {
        activity: { id: "a1", conversation: { id: "conv-1" } },
        sendActivity: mockSendActivity,
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

      expect(mockSendActivity).toHaveBeenCalledTimes(1);
      const call = mockSendActivity.mock.calls[0][0];
      expect(call.attachments).toHaveLength(1);

      // Verify notification was fired (fire-and-forget)
      expect(mockSendNotification).toHaveBeenCalledWith(
        expect.objectContaining({
          sessionId: "session-1",
          type: "permission",
        }),
      );
    });
  });

  describe("handleCardAction", () => {
    it("returns false for unknown verbs", async () => {
      const mockContext = {
        sendActivity: vi.fn().mockResolvedValue(undefined),
        updateActivity: vi.fn().mockResolvedValue(undefined),
      };

      const result = await handler.handleCardAction(
        mockContext as any,
        "unknown_verb",
        "session-1",
        "callback-key",
        "req-1",
      );
      expect(result).toBe(false);
    });

    it("returns true and sends expired message for unknown callback key", async () => {
      const mockSendActivity = vi.fn().mockResolvedValue(undefined);
      const mockContext = {
        sendActivity: mockSendActivity,
        updateActivity: vi.fn().mockResolvedValue(undefined),
      };

      const result = await handler.handleCardAction(
        mockContext as any,
        "allow",
        "session-1",
        "nonexistent-key",
        "req-1",
      );
      expect(result).toBe(true);
      expect(mockSendActivity).toHaveBeenCalledTimes(1);
      // Now sends an Adaptive Card instead of plain text
      const call = mockSendActivity.mock.calls[0][0];
      expect(call.attachments).toHaveLength(1);
    });

    it("resolves permission and updates card on valid action", async () => {
      const mockSendActivity = vi.fn().mockResolvedValue({ id: "activity-1" });
      const mockUpdateActivity = vi.fn().mockResolvedValue(undefined);
      const mockResolve = vi.fn();

      const session = {
        id: "session-1",
        permissionGate: { requestId: "req-1", resolve: mockResolve },
      };
      mockGetSession.mockReturnValue(session);

      // First, send a permission request to create a pending entry
      const sendContext = {
        activity: { id: "a1", conversation: { id: "conv-1" } },
        sendActivity: mockSendActivity,
      };
      await handler.sendPermissionRequest(
        session as any,
        {
          id: "req-1",
          description: "Allow?",
          options: [
            { id: "opt-allow", label: "Allow", isAllow: true },
            { id: "opt-deny", label: "Deny", isAllow: false },
          ],
        },
        sendContext as any,
      );

      // Extract the callback key from the ActionSet inside the Container
      const cardData = mockSendActivity.mock.calls[0][0].attachments[0].content;
      const container = cardData.body[0];
      const actionSet = container.items.find((i: any) => i.type === "ActionSet");
      const allowAction = actionSet.actions.find((a: any) => a.data.verb === "allow");
      const callbackKey = allowAction.data.callbackKey;

      // Now handle the card action
      const actionContext = {
        activity: { from: { id: "user-1", name: "Test User" } },
        sendActivity: vi.fn().mockResolvedValue(undefined),
        updateActivity: mockUpdateActivity,
      };

      const result = await handler.handleCardAction(
        actionContext as any,
        "allow",
        "session-1",
        callbackKey,
        "req-1",
      );

      expect(result).toBe(true);
      expect(mockResolve).toHaveBeenCalledWith("opt-allow");
    });
  });

  describe("evictStale", () => {
    it("does not crash with many pending entries", async () => {
      // Create 100+ pending entries to trigger eviction
      const mockSendActivity = vi.fn().mockResolvedValue({ id: "a1" });
      const mockContext = {
        activity: { id: "a1", conversation: { id: "c1" } },
        sendActivity: mockSendActivity,
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

      // Should not throw
      expect(mockSendActivity).toHaveBeenCalledTimes(110);
    });
  });
});
