import { randomUUID } from "node:crypto";
import { createMcpExpressApp } from "@modelcontextprotocol/sdk/server/express.js";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import express from "express";
import type { Request, Response } from "express";
import {
  authorizeHandler,
  dcrRegisterHandler,
  oauthAuthorizationServerHandler,
  oauthCallbackHandler,
  oauthProtectedResourceHandler,
  tokenExchangeHandler,
} from "./auth/endpoints.js";
import { requireBearerAuth } from "./auth/middleware.js";
import { GraphService } from "./graph/client.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatsTools } from "./tools/chats.js";
import { registerMeetingsTools } from "./tools/meetings.js";
import { registerMessagesTools } from "./tools/messages.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";

export function createConfiguredServer(userToken?: string): McpServer {
  const server = new McpServer({
    name: "mcp-teams-server",
    version: "1.0.0",
  });

  const graphService = GraphService.getInstance();

  registerAuthTools(server, graphService, userToken);
  registerUsersTools(server, graphService, userToken);
  registerTeamsTools(server, graphService, userToken);
  registerChatsTools(server, graphService, userToken);
  registerMeetingsTools(server, graphService, userToken);
  registerMessagesTools(server, graphService, userToken);

  return server;
}

export async function startStdioServer(): Promise<void> {
  const server = createConfiguredServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("MCP Teams Server started (stdio)");
}

export async function startHttpServer(oauth: boolean): Promise<void> {
  const port = Number.parseInt(process.env.TEAMS_MCP_PORT || "3000", 10);
  const app = createMcpExpressApp();
  app.use(express.urlencoded({ extended: true }));
  const transports: Record<string, StreamableHTTPServerTransport> = {};

  const mcpPostHandler = async (req: Request, res: Response) => {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    const method = req.body?.method || "unknown";
    const id = req.body?.id ?? null;
    console.log(`[MCP POST] method=${method} id=${id} session=${sessionId || "new"}`);

    try {
      let transport: StreamableHTTPServerTransport;

      if (sessionId && transports[sessionId]) {
        console.log(`[MCP POST] Reusing session: ${sessionId}`);
        transport = transports[sessionId];
      } else if (!sessionId && isInitializeRequest(req.body)) {
        console.log("[MCP POST] Initializing new session...");
        transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID(),
          onsessioninitialized: (sid) => {
            console.log(`[MCP POST] Session initialized: ${sid}`);
            transports[sid] = transport;
          },
        });
        transport.onclose = () => {
          const sid = transport.sessionId;
          if (sid && transports[sid]) {
            console.log(`[MCP POST] Session closed: ${sid}`);
            delete transports[sid];
          }
        };
        const sessionServer = createConfiguredServer(
          (req as unknown as Record<string, unknown>).auth
            ? ((req as unknown as Record<string, unknown>).auth as { token: string }).token
            : undefined
        );
        await sessionServer.connect(transport);
      } else {
        console.warn(`[MCP POST] Bad request - no valid session. sessionId=${sessionId}`);
        res.status(400).json({
          jsonrpc: "2.0",
          error: { code: -32000, message: "Bad Request: No valid session ID provided" },
          id: null,
        });
        return;
      }

      await transport.handleRequest(req, res, req.body);
      console.log(`[MCP POST] Handled successfully: method=${method} id=${id}`);
    } catch (error) {
      console.error(`[MCP POST] Error: method=${method} id=${id}`, error);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        });
      }
    }
  };

  const mcpGetHandler = async (req: Request, res: Response) => {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    console.log(`[MCP GET] SSE connection request, session=${sessionId || "none"}`);
    if (!sessionId || !transports[sessionId]) {
      console.warn(`[MCP GET] Invalid or missing session ID: ${sessionId}`);
      res.status(400).send("Invalid or missing session ID");
      return;
    }
    console.log(`[MCP GET] Opening SSE stream for session: ${sessionId}`);
    await transports[sessionId].handleRequest(req, res);
  };

  const mcpDeleteHandler = async (req: Request, res: Response) => {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    console.log(`[MCP DELETE] Termination request, session=${sessionId || "none"}`);
    if (!sessionId || !transports[sessionId]) {
      console.warn(`[MCP DELETE] Invalid or missing session ID: ${sessionId}`);
      res.status(400).send("Invalid or missing session ID");
      return;
    }
    try {
      await transports[sessionId].handleRequest(req, res);
      console.log(`[MCP DELETE] Session terminated: ${sessionId}`);
    } catch (error) {
      console.error(`[MCP DELETE] Error terminating session ${sessionId}:`, error);
      if (!res.headersSent) {
        res.status(500).send("Error processing session termination");
      }
    }
  };

  if (oauth) {
    const bearerAuth = requireBearerAuth();

    app.get("/.well-known/oauth-protected-resource", oauthProtectedResourceHandler());
    app.get("/.well-known/oauth-authorization-server", oauthAuthorizationServerHandler());
    app.get("/authorize", authorizeHandler(port));
    const callbackHandler = oauthCallbackHandler();
    app.get("/callback", callbackHandler);
    app.post("/register", dcrRegisterHandler());
    app.post("/token", tokenExchangeHandler(port));

    app.post("/mcp", bearerAuth, mcpPostHandler);
    app.get("/mcp", bearerAuth, mcpGetHandler);
    app.delete("/mcp", bearerAuth, mcpDeleteHandler);

    console.error(`OAuth enabled — MCP server on port ${port}`);
  } else {
    app.post("/mcp", mcpPostHandler);
    app.get("/mcp", mcpGetHandler);
    app.delete("/mcp", mcpDeleteHandler);

    console.error(`MCP Teams Server listening on port ${port}`);
  }

  app.listen(port, () => {
    console.error(`Server running on http://localhost:${port}`);
  });

  const shutdown = async () => {
    for (const sessionId in transports) {
      try {
        await transports[sessionId].close();
        delete transports[sessionId];
      } catch (error) {
        console.error(`Error closing transport for session ${sessionId}:`, error);
      }
    }
    process.exit(0);
  };

  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);
}
