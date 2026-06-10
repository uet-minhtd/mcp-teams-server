import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import {
  GetChatMessagesSchema,
  SendChatMessageSchema,
} from "../types.js";

export function registerChatsTools(server: McpServer, graphService: GraphService, userToken?: string): void {
  server.registerTool(
    "list_chats",
    {
      description: "List all chats (1-on-1 and group chats) the user is part of. Read-only, no data is modified.",
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
    }
  );

  server.registerTool(
    "get_chat_messages",
    {
      description:
        "Read messages from a specific Teams chat (1-on-1 or group). Read-only. " +
        "⚠️ This accesses private conversation data — only call when the user explicitly requests to view chat messages.",
      inputSchema: GetChatMessagesSchema.shape,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
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

  server.registerTool(
    "send_chat_message",
    {
      description:
        "⚠️ WRITE ACTION: Send a message to a Teams chat on behalf of the user. " +
        "This action is irreversible — the message will be delivered immediately to all chat participants. " +
        "Always confirm the exact message content and target chat with the user before calling this tool.",
      inputSchema: SendChatMessageSchema.shape,
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
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
