import { describe, it, expect, vi, beforeEach } from "vitest";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerAuthTools } from "../tools/auth.js";
import { registerTeamsTools } from "../tools/teams.js";
import type { GraphService } from "../graph/client.js";

type ToolHandler = (args: unknown) => Promise<{
  content: Array<{ type: string; text: string }>;
  isError?: boolean;
}>;

function createMockServer(): McpServer & { tools: Map<string, ToolHandler> } {
  const tools = new Map<string, ToolHandler>();
  return {
    tools,
    tool: vi
      .fn()
      .mockImplementation(
        (name: string, _desc: string, _schema: unknown, handler: ToolHandler) => {
          tools.set(name, handler);
        }
      ),
    prompt: vi.fn(),
    resource: vi.fn(),
  } as unknown as McpServer & { tools: Map<string, ToolHandler> };
}

function createMockGraphService(
  overrides: Partial<GraphService> = {}
): GraphService {
  return {
    getClient: vi.fn(),
    getAuthStatus: vi.fn(),
    validateToken: vi.fn(),
    ...overrides,
  } as unknown as GraphService;
}

describe("Tool Registration", () => {
  describe("registerAuthTools", () => {
    let server: ReturnType<typeof createMockServer>;
    let graphService: ReturnType<typeof createMockGraphService>;

    beforeEach(() => {
      server = createMockServer();
      graphService = createMockGraphService();
      registerAuthTools(server as unknown as McpServer, graphService, undefined);
    });

    it("registers auth_status tool", () => {
      expect(server.tools.has("auth_status")).toBe(true);
    });

    it("auth_status returns authenticated user info", async () => {
      vi.mocked(graphService.getAuthStatus).mockResolvedValue({
        isAuthenticated: true,
        userPrincipalName: "user@test.com",
        displayName: "Test User",
      });

      const handler = server.tools.get("auth_status")!;
      const result = await handler({});

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.authenticated).toBe(true);
      expect(parsed.userPrincipalName).toBe("user@test.com");
      expect(parsed.displayName).toBe("Test User");
    });

    it("auth_status returns not authenticated", async () => {
      vi.mocked(graphService.getAuthStatus).mockResolvedValue({
        isAuthenticated: false,
      });

      const handler = server.tools.get("auth_status")!;
      const result = await handler({});

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.authenticated).toBe(false);
    });

    it("auth_status returns error on exception", async () => {
      vi.mocked(graphService.getAuthStatus).mockRejectedValue(
        new Error("Connection failed")
      );

      const handler = server.tools.get("auth_status")!;
      const result = await handler({});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Connection failed");
    });
  });

  describe("registerTeamsTools", () => {
    let server: ReturnType<typeof createMockServer>;
    let graphService: ReturnType<typeof createMockGraphService>;

    beforeEach(() => {
      server = createMockServer();
      graphService = createMockGraphService();
      registerTeamsTools(server as unknown as McpServer, graphService, undefined);
    });

    it("registers 4 tools", () => {
      expect(server.tools.has("list_teams")).toBe(true);
      expect(server.tools.has("list_channels")).toBe(true);
      expect(server.tools.has("get_channel_messages")).toBe(true);
      expect(server.tools.has("send_channel_message")).toBe(true);
    });

    it("list_teams returns results on success", async () => {
      const mockClient = {
        api: vi.fn().mockReturnThis(),
        get: vi.fn().mockResolvedValue({ value: [{ id: "team1" }] }),
      };
      vi.mocked(graphService.getClient).mockResolvedValue(mockClient as any);

      const handler = server.tools.get("list_teams")!;
      const result = await handler({});

      expect(result.content[0].text).toContain("team1");
      expect(result.isError).toBeUndefined();
    });

    it("list_teams returns error on failure", async () => {
      vi.mocked(graphService.getClient).mockRejectedValue(
        new Error("Graph API error")
      );

      const handler = server.tools.get("list_teams")!;
      const result = await handler({});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Graph API error");
    });

    it("send_channel_message returns success on post", async () => {
      const mockClient = {
        api: vi.fn().mockReturnThis(),
        post: vi.fn().mockResolvedValue({ id: "msg-123" }),
      };
      vi.mocked(graphService.getClient).mockResolvedValue(mockClient as any);

      const handler = server.tools.get("send_channel_message")!;
      const result = await handler({
        teamId: "t1",
        channelId: "c1",
        message: "Hello",
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.messageId).toBe("msg-123");
    });

    it("send_channel_message returns error on failure", async () => {
      vi.mocked(graphService.getClient).mockRejectedValue(
        new Error("Send failed")
      );

      const handler = server.tools.get("send_channel_message")!;
      const result = await handler({
        teamId: "t1",
        channelId: "c1",
        message: "Hello",
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Send failed");
    });

    it("get_channel_messages passes limit to API", async () => {
      const mockApi = vi.fn().mockReturnThis();
      const mockTop = vi.fn().mockReturnThis();
      const mockGet = vi.fn().mockResolvedValue({ value: [] });
      const mockClient = {
        api: mockApi,
        top: mockTop,
        get: mockGet,
      };
      vi.mocked(graphService.getClient).mockResolvedValue(mockClient as any);

      const handler = server.tools.get("get_channel_messages")!;
      await handler({ teamId: "t1", channelId: "c1", limit: 10 });

      expect(mockApi).toHaveBeenCalledWith("/teams/t1/channels/c1/messages");
      expect(mockTop).toHaveBeenCalledWith(10);
    });
  });
});
