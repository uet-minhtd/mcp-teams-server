import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import {
  ListChannelsSchema,
  GetChannelMessagesSchema,
  SendChannelMessageSchema,
} from "../types.js";

export function registerTeamsTools(server: McpServer, graphService: GraphService, userToken?: string): void {
  server.tool("list_teams", "List all teams the user is a member of", {}, async () => {
    try {
      const client = await graphService.getClient(userToken);
      const result = await client.api("/me/joinedTeams").get();
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        content: [{ type: "text" as const, text: `Failed to list teams: ${message}` }],
        isError: true,
      };
    }
  });

  server.tool(
    "list_channels",
    "List channels in a specific team",
    ListChannelsSchema.shape,
    async ({ teamId }) => {
      try {
        const client = await graphService.getClient(userToken);
        const result = await client.api(`/teams/${teamId}/channels`).get();
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            { type: "text" as const, text: `Failed to list channels: ${message}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "get_channel_messages",
    "Get messages from a team channel",
    GetChannelMessagesSchema.shape,
    async ({ teamId, channelId, limit }) => {
      try {
        const client = await graphService.getClient(userToken);
        const result = await client
          .api(`/teams/${teamId}/channels/${channelId}/messages`)
          .top(limit)
          .get();
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            { type: "text" as const, text: `Failed to get messages: ${message}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "send_channel_message",
    "Send a message to a team channel",
    SendChannelMessageSchema.shape,
    async ({ teamId, channelId, message }) => {
      try {
        const client = await graphService.getClient(userToken);
        const result = await client
          .api(`/teams/${teamId}/channels/${channelId}/messages`)
          .post({
            body: {
              content: message,
              contentType: "text",
            },
          });
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({ success: true, messageId: result.id }, null, 2),
            },
          ],
        };
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        return {
          content: [
            { type: "text" as const, text: `Failed to send message: ${msg}` },
          ],
          isError: true,
        };
      }
    }
  );
}
