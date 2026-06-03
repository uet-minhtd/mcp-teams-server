import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import {
  GetChatMessagesSchema,
  SendChatMessageSchema,
} from "../types.js";

export function registerChatsTools(server: McpServer, graphService: GraphService, userToken?: string): void {
  server.tool("list_chats", "List all chats the user is part of", {}, async () => {
    try {
      const client = await graphService.getClient(userToken);
      const result = await client.api("/me/chats").get();
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        content: [{ type: "text" as const, text: `Failed to list chats: ${message}` }],
        isError: true,
      };
    }
  });

  server.tool(
    "get_chat_messages",
    "Get messages from a specific chat",
    GetChatMessagesSchema.shape,
    async ({ chatId, limit }) => {
      try {
        const client = await graphService.getClient(userToken);
        const result = await client
          .api(`/me/chats/${chatId}/messages`)
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
    "send_chat_message",
    "Send a message to a chat",
    SendChatMessageSchema.shape,
    async ({ chatId, message }) => {
      try {
        const client = await graphService.getClient(userToken);
        const result = await client.api(`/chats/${chatId}/messages`).post({
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
