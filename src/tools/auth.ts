import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";

export function registerAuthTools(server: McpServer, graphService: GraphService, userToken?: string): void {
  server.tool(
    "auth_status",
    "Check current authentication status with Microsoft Graph",
    {},
    async () => {
      try {
        const status = userToken
          ? await graphService.validateToken(userToken)
          : await graphService.getAuthStatus();
        if (status.isAuthenticated) {
          return {
            content: [
              {
                type: "text" as const,
                text: JSON.stringify(
                  {
                    authenticated: true,
                    userPrincipalName: status.userPrincipalName,
                    displayName: status.displayName,
                  },
                  null,
                  2
                ),
              },
            ],
          };
        }
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({ authenticated: false }, null, 2),
            },
          ],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [{ type: "text" as const, text: `Auth check failed: ${message}` }],
          isError: true,
        };
      }
    }
  );
}
