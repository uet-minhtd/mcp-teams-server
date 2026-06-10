import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import { SearchMessagesSchema } from "../types.js";

export function registerMessagesTools(
  server: McpServer,
  graphService: GraphService,
  userToken?: string
): void {
  server.registerTool(
    "search_messages",
    {
      description:
        "Search across all Teams messages (channels and chats) using Microsoft Search API. Read-only. " +
        "⚠️ May return content from private chats and channels — only call when the user explicitly requests a search.",
      inputSchema: SearchMessagesSchema.shape,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ query, limit }) => {
      try {
        const client = await graphService.getClient(userToken);
        const result = await client
          .api("/search/query")
          .post({
            requests: [
              {
                entityTypes: ["chatMessage"],
                query: { queryString: query },
                size: limit,
              },
            ],
          });
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            { type: "text" as const, text: `Search failed: ${message}` },
          ],
          isError: true,
        };
      }
    }
  );
}
