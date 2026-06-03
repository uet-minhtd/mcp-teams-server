import { describe, it, expect, vi, beforeEach } from "vitest";
import { GraphService } from "../graph/client.js";

vi.mock("@microsoft/microsoft-graph-client", () => ({
  Client: {
    initWithMiddleware: vi.fn().mockImplementation((config) => ({
      config,
      api: vi.fn().mockReturnThis(),
      get: vi.fn(),
    })),
  },
}));

const { Client } = await import("@microsoft/microsoft-graph-client");

describe("GraphService", () => {
  let graphService: GraphService;

  beforeEach(() => {
    vi.clearAllMocks();
    GraphService["instance"] = undefined as unknown as GraphService;
    graphService = GraphService.getInstance();
  });

  describe("getInstance", () => {
    it("returns singleton instance", () => {
      const a = GraphService.getInstance();
      const b = GraphService.getInstance();
      expect(a).toBe(b);
    });
  });

  describe("getClient", () => {
    it("initializes with env credentials", async () => {
      const client = await graphService.getClient();
      expect(client).toBeDefined();
      expect(Client.initWithMiddleware).toHaveBeenCalled();
    });

    it("returns same client on subsequent calls", async () => {
      const client1 = await graphService.getClient();
      const client2 = await graphService.getClient();
      expect(client1).toBeDefined();
      expect(client2).toBeDefined();
    });
  });

  describe("getAuthStatus", () => {
    it("returns authenticated when Graph API responds", async () => {
      const mockClient = {
        api: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          userPrincipalName: "user@test.com",
          displayName: "Test User",
        }),
      };
      vi.mocked(Client.initWithMiddleware).mockReturnValueOnce(mockClient as any);
      GraphService["instance"] = undefined as unknown as GraphService;
      graphService = GraphService.getInstance();

      const status = await graphService.getAuthStatus();
      expect(status.isAuthenticated).toBe(true);
      expect(status.userPrincipalName).toBe("user@test.com");
      expect(status.displayName).toBe("Test User");
    });

    it("returns not authenticated on error", async () => {
      const mockClient = {
        api: vi.fn().mockReturnThis(),
        get: vi.fn().mockRejectedValue(new Error("Network error")),
      };
      vi.mocked(Client.initWithMiddleware).mockReturnValueOnce(mockClient as any);
      GraphService["instance"] = undefined as unknown as GraphService;
      graphService = GraphService.getInstance();

      const status = await graphService.getAuthStatus();
      expect(status.isAuthenticated).toBe(false);
    });
  });

  describe("validateToken", () => {
    it("returns authenticated for valid token", async () => {
      const mockClient = {
        api: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({
          userPrincipalName: "user@test.com",
          displayName: "Valid User",
        }),
      };
      vi.mocked(Client.initWithMiddleware).mockReturnValueOnce(mockClient as any);

      const status = await graphService.validateToken("valid-token");
      expect(status.isAuthenticated).toBe(true);
      expect(status.userPrincipalName).toBe("user@test.com");
    });

    it("returns not authenticated for invalid token", async () => {
      const mockClient = {
        api: vi.fn().mockReturnThis(),
        get: vi.fn().mockRejectedValue(new Error("401 Unauthorized")),
      };
      vi.mocked(Client.initWithMiddleware).mockReturnValueOnce(mockClient as any);

      const status = await graphService.validateToken("invalid-token");
      expect(status.isAuthenticated).toBe(false);
    });
  });
});
