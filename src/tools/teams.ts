import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import {
  ListChannelsSchema,
  GetChannelMessagesSchema,
  SendChannelMessageSchema,
} from "../types.js";

export function registerTeamsTools(server: McpServer, graphService: GraphService, userToken?: string): void {
  server.registerTool(
    "list_teams",
    {
      description: "List all teams the user is a member of. Read-only, no data is modified.",
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async () => {
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
    }
  );

  server.registerTool(
    "list_channels",
    {
      description: "List channels in a specific team. Read-only, no data is modified.",
      inputSchema: ListChannelsSchema.shape,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
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

  server.registerTool(
    "get_channel_messages",
    {
      description:
        "Read messages from a Teams channel. Read-only. " +
        "⚠️ This accesses potentially private conversation data — only call when the user explicitly requests to view channel messages.",
      inputSchema: GetChannelMessagesSchema.shape,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
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

  server.registerTool(
    "send_channel_message",
    {
      description:
        "⚠️ WRITE ACTION: Send a message to a Teams channel on behalf of the user. " +
        "This action is irreversible — the message will be visible to all channel members. " +
        "Always confirm the exact message content and target channel with the user before calling this tool.",
      inputSchema: SendChannelMessageSchema.shape,
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
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
