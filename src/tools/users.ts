import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import { SearchUsersSchema } from "../types.js";

export function registerUsersTools(server: McpServer, graphService: GraphService, userToken?: string): void {
  server.tool(
    "get_current_user",
    "Get the authenticated user's profile information",
    {},
    async () => {
      try {
        const client = await graphService.getClient(userToken);
        const user = await client.api("/me").get();
        return {
          content: [{ type: "text" as const, text: JSON.stringify(user, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [{ type: "text" as const, text: `Failed to get user: ${message}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "search_users",
    "Search for users by name or email",
    SearchUsersSchema.shape,
    async ({ query }) => {
      try {
        const client = await graphService.getClient(userToken);
        const escaped = query.replace(/'/g, "''");
        const result = await client
          .api("/users")
          .filter(
            `startswith(displayName,'${escaped}') or startswith(userPrincipalName,'${escaped}')`
          )
          .get();
        return {
          content: [
            { type: "text" as const, text: JSON.stringify(result, null, 2) },
          ],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [{ type: "text" as const, text: `Search failed: ${message}` }],
          isError: true,
        };
      }
    }
  );
}
