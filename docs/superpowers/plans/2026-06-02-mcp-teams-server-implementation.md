# MCP Teams Server Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a custom MCP server with OAuth 2.1 authentication that exposes 16 tools for Microsoft Teams (users, teams, channels, chats, meetings, messages) via Microsoft Graph API.

**Architecture:** Express HTTP server hosts MCP Streamable HTTP transport. Two auth layers: MCP OAuth 2.1 (client↔server, validates via Graph `/me`) and Graph ClientSecretCredential (server→Graph). Tools are registered as MCP primitives on the McpServer instance.

**Tech Stack:** Node.js 20+, TypeScript, `@modelcontextprotocol/sdk`, `@microsoft/microsoft-graph-client`, `@azure/identity`, `express`, `dotenv`, `tsx`

---

## File Structure

```
mcp-teams-from-scratch/
├── .env.example
├── package.json
├── tsconfig.json
├── src/
│   ├── index.ts              # CLI entry: parse args, start stdio or http
│   ├── server.ts             # McpServer factory, createHttpServer, createStdioServer
│   ├── types.ts              # Tool input schemas (zod), response types
│   ├── auth/
│   │   ├── middleware.ts     # requireBearerAuth express middleware
│   │   └── endpoints.ts     # OAuth endpoints: metadata, authorize, callback, token, dcr
│   ├── graph/
│   │   └── client.ts         # GraphService singleton with ClientSecretCredential
│   └── tools/
│       ├── auth.ts           # registerAuthTools
│       ├── users.ts          # registerUsersTools
│       ├── teams.ts          # registerTeamsTools
│       ├── chats.ts          # registerChatsTools
│       ├── meetings.ts       # registerMeetingsTools
│       └── messages.ts       # registerMessagesTools
```

---

### Task 1: Project Scaffolding

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\package.json`
- Create: `D:\2026\mcp-teams-from-scratch\tsconfig.json`
- Create: `D:\2026\mcp-teams-from-scratch\.env.example`

- [ ] **Step 1: Write package.json**

```json
{
  "name": "mcp-teams-from-scratch",
  "version": "1.0.0",
  "description": "Custom MCP server for Microsoft Teams via Microsoft Graph",
  "type": "module",
  "main": "dist/index.js",
  "scripts": {
    "build": "tsc",
    "dev": "tsx src/index.ts",
    "start": "node dist/index.js"
  },
  "dependencies": {
    "@azure/identity": "^4.13.1",
    "@microsoft/microsoft-graph-client": "^3.0.7",
    "@modelcontextprotocol/sdk": "^1.28.0",
    "dotenv": "^17.2.3",
    "express": "^5.2.1",
    "zod": "^4.3.6"
  },
  "devDependencies": {
    "@types/express": "^5.0.6",
    "@types/node": "^25.5.0",
    "tsx": "^4.21.0",
    "typescript": "^5.9.3"
  },
  "engines": {
    "node": ">=20"
  }
}
```

- [ ] **Step 2: Write tsconfig.json**

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "outDir": "dist",
    "rootDir": "src",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true,
    "declaration": true,
    "sourceMap": true
  },
  "include": ["src/**/*.ts"],
  "exclude": ["node_modules", "dist"]
}
```

- [ ] **Step 3: Write .env.example**

```
AZURE_CLIENT_ID=your-client-id-here
AZURE_CLIENT_SECRET=your-client-secret-here
AZURE_TENANT_ID=your-tenant-id-here
TEAMS_MCP_PORT=3000
```

- [ ] **Step 4: Install dependencies**

```bash
cd D:\2026\mcp-teams-from-scratch; npm install
```

- [ ] **Step 5: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: "error TS18003: No inputs were found in config file" (no src files yet — acceptable)

---

### Task 2: Shared Types

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\types.ts`

- [ ] **Step 1: Write types.ts**

```typescript
import { z } from "zod";

// Tool input schemas

export const SearchUsersSchema = z.object({
  query: z.string().describe("Name or email to search for"),
});

export const GetUserSchema = z.object({
  userId: z.string().describe("User ID or userPrincipalName"),
});

export const ListChannelsSchema = z.object({
  teamId: z.string().describe("Team ID"),
});

export const GetChannelMessagesSchema = z.object({
  teamId: z.string().describe("Team ID"),
  channelId: z.string().describe("Channel ID"),
  limit: z.number().optional().default(20).describe("Max messages to return"),
});

export const SendChannelMessageSchema = z.object({
  teamId: z.string().describe("Team ID"),
  channelId: z.string().describe("Channel ID"),
  message: z.string().describe("Message content"),
});

export const GetChatMessagesSchema = z.object({
  chatId: z.string().describe("Chat ID"),
  limit: z.number().optional().default(20).describe("Max messages to return"),
});

export const SendChatMessageSchema = z.object({
  chatId: z.string().describe("Chat ID"),
  message: z.string().describe("Message content"),
});

export const ListMeetingsSchema = z.object({
  limit: z.number().optional().default(20).describe("Max meetings to return"),
});

export const GetMeetingSchema = z.object({
  meetingId: z.string().describe("Meeting ID"),
});

export const GetMeetingAttendanceSchema = z.object({
  meetingId: z.string().describe("Meeting ID"),
});

export const GetMeetingTranscriptsSchema = z.object({
  meetingId: z.string().describe("Meeting ID"),
});

export const SearchMessagesSchema = z.object({
  query: z.string().describe("Search query (KQL syntax)"),
  limit: z.number().optional().default(20).describe("Max results to return"),
});

// Auth context attached to request by middleware

export interface AuthContext {
  token: string;
  userId: string;
  displayName: string;
  userPrincipalName: string;
}
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors.

---

### Task 3: Graph Client

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\graph\client.ts`

- [ ] **Step 1: Write graph/client.ts**

```typescript
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export interface AuthStatus {
  isAuthenticated: boolean;
  userPrincipalName?: string;
  displayName?: string;
}

export class GraphService {
  private static instance: GraphService;
  private client: Client | undefined;
  private credential: ClientSecretCredential | undefined;
  private initialized = false;

  static getInstance(): GraphService {
    if (!GraphService.instance) {
      GraphService.instance = new GraphService();
    }
    return GraphService.instance;
  }

  private initialize(): void {
    if (this.initialized) return;

    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    const tenantId = process.env.AZURE_TENANT_ID;

    if (!clientId || !clientSecret || !tenantId) {
      throw new Error(
        "Missing AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, or AZURE_TENANT_ID in environment"
      );
    }

    this.credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

    this.client = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          if (!this.credential) throw new Error("Credential not initialized");
          const result = await this.credential.getToken(
            "https://graph.microsoft.com/.default"
          );
          if (!result) throw new Error("Failed to acquire token");
          return result.token;
        },
      },
      baseUrl: GRAPH_BASE_URL,
    });

    this.initialized = true;
  }

  async getClient(): Promise<Client> {
    this.initialize();
    if (!this.client) throw new Error("Graph client not initialized");
    return this.client;
  }

  async getAuthStatus(): Promise<AuthStatus> {
    try {
      const client = await this.getClient();
      const me = await client.api("/me").get();
      return {
        isAuthenticated: true,
        userPrincipalName: me.userPrincipalName as string | undefined,
        displayName: me.displayName as string | undefined,
      };
    } catch {
      return { isAuthenticated: false };
    }
  }

  async validateToken(token: string): Promise<AuthStatus> {
    try {
      const tempClient = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => token,
        },
        baseUrl: GRAPH_BASE_URL,
      });
      const me = await tempClient.api("/me").get();
      return {
        isAuthenticated: true,
        userPrincipalName: me.userPrincipalName as string | undefined,
        displayName: me.displayName as string | undefined,
      };
    } catch {
      return { isAuthenticated: false };
    }
  }
}
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors.

---

### Task 4: Auth Middleware

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\auth\middleware.ts`

- [ ] **Step 1: Write auth/middleware.ts**

```typescript
import type { NextFunction, Request, Response } from "express";
import { GraphService } from "../graph/client.js";
import type { AuthContext } from "../types.js";

function buildAuthHeader(req: Request): string {
  const host = req.get("host") || "localhost";
  const proto = req.protocol;
  const baseUrl = `${proto}://${host}`;
  return `Bearer resource_metadata="${baseUrl}/.well-known/oauth-protected-resource",authorization_servers="${baseUrl}"`;
}

export function requireBearerAuth() {
  return async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    const authHeader = req.headers.authorization;

    if (!authHeader?.startsWith("Bearer ")) {
      res.setHeader("WWW-Authenticate", buildAuthHeader(req));
      res.status(401).json({
        jsonrpc: "2.0",
        error: {
          code: -32001,
          message: "Unauthorized: Bearer token required",
        },
        id: null,
      });
      return;
    }

    const token = authHeader.slice(7);
    const graphService = GraphService.getInstance();
    const status = await graphService.validateToken(token);

    if (!status.isAuthenticated) {
      res.setHeader("WWW-Authenticate", buildAuthHeader(req));
      res.status(401).json({
        jsonrpc: "2.0",
        error: {
          code: -32001,
          message: "Unauthorized: Invalid or expired token",
        },
        id: null,
      });
      return;
    }

    (req as unknown as Record<string, unknown>).auth = {
      token,
      userId: status.userPrincipalName,
      displayName: status.displayName ?? "",
      userPrincipalName: status.userPrincipalName ?? "",
    } satisfies AuthContext;

    next();
  };
}
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors.

---

### Task 5: OAuth Endpoints

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\auth\endpoints.ts`

- [ ] **Step 1: Write auth/endpoints.ts**

```typescript
import type { Request, Response } from "express";

const MICROSOFT_AUTHORITY = "https://login.microsoftonline.com/common";
const MS_AUTH_ENDPOINT = `${MICROSOFT_AUTHORITY}/oauth2/v2.0/authorize`;
const MS_TOKEN_ENDPOINT = `${MICROSOFT_AUTHORITY}/oauth2/v2.0/token`;

const OIDC_SCOPES = ["openid", "profile", "offline_access"];

const DCR_SCOPES = [
  ...OIDC_SCOPES,
  "User.Read",
  "User.ReadBasic.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "ChannelMessage.Send",
  "Chat.Read",
  "Chat.ReadWrite",
  "TeamMember.Read.All",
  "OnlineMeetings.Read",
  "OnlineMeetingArtifact.Read.All",
];

const MICROSOFT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

const CALLBACK_PROXY_TIMEOUT_MS = 5 * 60 * 1000;
const pendingCallbacks = new Map<string, { redirectUri: string; timestamp: number }>();

setInterval(() => {
  const now = Date.now();
  for (const [key, value] of pendingCallbacks) {
    if (now - value.timestamp > CALLBACK_PROXY_TIMEOUT_MS) {
      pendingCallbacks.delete(key);
    }
  }
}, 60_000);

export function oauthProtectedResourceHandler() {
  return (_req: Request, res: Response) => {
    const host = _req.get("host") || "localhost";
    const proto = _req.protocol;
    const baseUrl = `${proto}://${host}`;
    res.json({ authorization_servers: [baseUrl] });
  };
}

export function oauthAuthorizationServerHandler() {
  return (req: Request, res: Response) => {
    const host = req.get("host") || "localhost";
    const proto = req.protocol;
    const baseUrl = `${proto}://${host}`;
    res.json({
      issuer: baseUrl,
      authorization_endpoint: `${baseUrl}/authorize`,
      token_endpoint: `${baseUrl}/token`,
      registration_endpoint: `${baseUrl}/register`,
      response_types_supported: ["code"],
      code_challenge_methods_supported: ["S256"],
      response_modes_supported: ["query"],
      scopes_supported: DCR_SCOPES,
      grant_types_supported: ["authorization_code", "refresh_token"],
      token_endpoint_auth_methods_supported: ["none", "client_secret_post"],
      jwks_uri: `${MICROSOFT_AUTHORITY}/discovery/v2.0/keys`,
    });
  };
}

export function authorizeHandler(port: number) {
  return (req: Request, res: Response) => {
    const redirectUrl = new URL(MS_AUTH_ENDPOINT);
    for (const [key, value] of Object.entries(req.query)) {
      if (typeof value === "string") {
        redirectUrl.searchParams.set(key, value);
      }
    }
    if (!redirectUrl.searchParams.has("scope")) {
      redirectUrl.searchParams.set("scope", DCR_SCOPES.join(" "));
    }
    const ru = redirectUrl.searchParams.get("redirect_uri");
    const state = redirectUrl.searchParams.get("state");
    if (ru && state) {
      pendingCallbacks.set(state, { redirectUri: ru, timestamp: Date.now() });
      const proxyRu = `http://localhost:${port}`;
      redirectUrl.searchParams.set("redirect_uri", proxyRu);
    }
    res.redirect(302, redirectUrl.toString());
  };
}

export function oauthCallbackHandler() {
  return (req: Request, res: Response) => {
    const state = req.query.state as string | undefined;
    const code = req.query.code as string | undefined;
    if (!state || !code) {
      res.status(400).send("Missing state or code parameter");
      return;
    }
    const pending = pendingCallbacks.get(state);
    if (!pending) {
      res.status(400).send("Invalid or expired state");
      return;
    }
    pendingCallbacks.delete(state);
    const redirectTo = new URL(pending.redirectUri);
    redirectTo.searchParams.set("code", code);
    redirectTo.searchParams.set("state", state);
    res.redirect(302, redirectTo.toString());
  };
}

export function dcrRegisterHandler() {
  return (_req: Request, res: Response) => {
    res.json({
      client_id: MICROSOFT_CLIENT_ID,
      client_id_issued_at: Math.floor(Date.now() / 1000),
      client_secret_expires_at: 0,
      grant_types: ["authorization_code", "refresh_token"],
      response_types: ["code"],
      redirect_uris: ["http://localhost", "http://127.0.0.1"],
      scope: DCR_SCOPES.join(" "),
    });
  };
}

export function tokenExchangeHandler(port: number) {
  return async (req: Request, res: Response) => {
    try {
      const body = new URLSearchParams();
      for (const [key, value] of Object.entries(req.body)) {
        if (typeof value === "string") {
          body.set(key, value);
        }
      }
      body.set("redirect_uri", `http://localhost:${port}`);
      if (!body.has("scope")) {
        body.set("scope", DCR_SCOPES.join(" "));
      }

      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 15_000);
      const response = await fetch(MS_TOKEN_ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body,
        signal: controller.signal,
      });
      clearTimeout(timeout);

      const data = (await response.json()) as Record<string, unknown>;
      res.status(response.status).json(data);
    } catch (error) {
      console.error("Token exchange failed:", error);
      res.status(500).json({
        error: "token_exchange_failed",
        error_description: "Failed to exchange authorization code",
      });
    }
  };
}
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors.

---

### Task 6: MCP Server + HTTP Transport

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\server.ts`

- [ ] **Step 1: Write server.ts**

```typescript
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

export function createConfiguredServer(): McpServer {
  const server = new McpServer({
    name: "mcp-teams-server",
    version: "1.0.0",
  });

  const graphService = GraphService.getInstance();

  registerAuthTools(server, graphService);
  registerUsersTools(server, graphService);
  registerTeamsTools(server, graphService);
  registerChatsTools(server, graphService);
  registerMeetingsTools(server, graphService);
  registerMessagesTools(server, graphService);

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
  const bearerAuth = oauth ? requireBearerAuth() : undefined;
  const graphService = GraphService.getInstance();

  const mcpPostHandler = async (req: Request, res: Response) => {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;

    try {
      let transport: StreamableHTTPServerTransport;

      if (sessionId && transports[sessionId]) {
        transport = transports[sessionId];
      } else if (!sessionId && isInitializeRequest(req.body)) {
        transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID(),
          onsessioninitialized: (sid) => {
            transports[sid] = transport;
          },
        });
        transport.onclose = () => {
          const sid = transport.sessionId;
          if (sid && transports[sid]) {
            delete transports[sid];
          }
        };
        const sessionServer = createConfiguredServer();
        await (sessionServer as any).connect(transport);
      } else {
        res.status(400).json({
          jsonrpc: "2.0",
          error: { code: -32000, message: "Bad Request: No valid session ID provided" },
          id: null,
        });
        return;
      }

      await transport.handleRequest(req, res, req.body);
    } catch (error) {
      console.error("Error handling MCP request:", error);
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
    if (!sessionId || !transports[sessionId]) {
      res.status(400).send("Invalid or missing session ID");
      return;
    }
    await transports[sessionId].handleRequest(req, res);
  };

  const mcpDeleteHandler = async (req: Request, res: Response) => {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    if (!sessionId || !transports[sessionId]) {
      res.status(400).send("Invalid or missing session ID");
      return;
    }
    try {
      await transports[sessionId].handleRequest(req, res);
    } catch (error) {
      console.error("Error handling session termination:", error);
      if (!res.headersSent) {
        res.status(500).send("Error processing session termination");
      }
    }
  };

  if (oauth && bearerAuth) {
    app.get("/.well-known/oauth-protected-resource", oauthProtectedResourceHandler());
    app.get("/.well-known/oauth-authorization-server", oauthAuthorizationServerHandler());
    app.get("/authorize", authorizeHandler(port));
    const callbackHandler = oauthCallbackHandler();
    app.use((req, _res, next) => {
      if (req.path === "/" && req.query.code && req.query.state) {
        callbackHandler(req, _res);
      } else {
        next();
      }
    });
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

    console.error(`MCP Team Server listening on port ${port}`);
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
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors (will fail until tool files exist, expected — continue to next tasks).

---

### Task 7: Auth Tool

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\tools\auth.ts`

- [ ] **Step 1: Write tools/auth.ts**

```typescript
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";

export function registerAuthTools(server: McpServer, graphService: GraphService): void {
  server.tool(
    "auth_status",
    "Check current authentication status with Microsoft Graph",
    {},
    async () => {
      try {
        const status = await graphService.getAuthStatus();
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
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors (other tool files still missing — OK).

---

### Task 8: Users Tools

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\tools\users.ts`

- [ ] **Step 1: Write tools/users.ts**

```typescript
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import { SearchUsersSchema, GetUserSchema } from "../types.js";

export function registerUsersTools(server: McpServer, graphService: GraphService): void {
  server.tool(
    "get_current_user",
    "Get the authenticated user's profile information",
    {},
    async () => {
      try {
        const client = await graphService.getClient();
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
        const client = await graphService.getClient();
        const result = await client
          .api("/users")
          .filter(
            `startswith(displayName,'${query}') or startswith(userPrincipalName,'${query}')`
          )
          .top(20)
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
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors.

---

### Task 9: Teams Tools

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\tools\teams.ts`

- [ ] **Step 1: Write tools/teams.ts**

```typescript
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import {
  ListChannelsSchema,
  GetChannelMessagesSchema,
  SendChannelMessageSchema,
} from "../types.js";

export function registerTeamsTools(server: McpServer, graphService: GraphService): void {
  server.tool("list_teams", "List all teams the user is a member of", {}, async () => {
    try {
      const client = await graphService.getClient();
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
        const client = await graphService.getClient();
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
        const client = await graphService.getClient();
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
        const client = await graphService.getClient();
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
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            { type: "text" as const, text: `Failed to send message: ${message}` },
          ],
          isError: true,
        };
      }
    }
  );
}
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors.

---

### Task 10: Chats Tools

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\tools\chats.ts`

- [ ] **Step 1: Write tools/chats.ts**

```typescript
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import {
  GetChatMessagesSchema,
  SendChatMessageSchema,
} from "../types.js";

export function registerChatsTools(server: McpServer, graphService: GraphService): void {
  server.tool("list_chats", "List all chats the user is part of", {}, async () => {
    try {
      const client = await graphService.getClient();
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
        const client = await graphService.getClient();
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
        const client = await graphService.getClient();
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
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors.

---

### Task 11: Meetings Tools

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\tools\meetings.ts`

- [ ] **Step 1: Write tools/meetings.ts**

```typescript
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import {
  ListMeetingsSchema,
  GetMeetingSchema,
  GetMeetingAttendanceSchema,
  GetMeetingTranscriptsSchema,
} from "../types.js";

export function registerMeetingsTools(
  server: McpServer,
  graphService: GraphService
): void {
  server.tool(
    "list_meetings",
    "List online meetings (upcoming and past)",
    ListMeetingsSchema.shape,
    async ({ limit }) => {
      try {
        const client = await graphService.getClient();
        const result = await client
          .api("/me/onlineMeetings")
          .top(limit)
          .get();
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            { type: "text" as const, text: `Failed to list meetings: ${message}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "get_meeting",
    "Get details of a specific online meeting",
    GetMeetingSchema.shape,
    async ({ meetingId }) => {
      try {
        const client = await graphService.getClient();
        const result = await client
          .api(`/me/onlineMeetings/${meetingId}`)
          .get();
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            { type: "text" as const, text: `Failed to get meeting: ${message}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "get_meeting_attendance",
    "Get attendance reports and records for a meeting",
    GetMeetingAttendanceSchema.shape,
    async ({ meetingId }) => {
      try {
        const client = await graphService.getClient();
        const reports = await client
          .api(`/me/onlineMeetings/${meetingId}/attendanceReports`)
          .get();

        const reportsList = reports.value as Array<{ id: string }>;
        const detailed = [];

        for (const report of reportsList) {
          const records = await client
            .api(
              `/me/onlineMeetings/${meetingId}/attendanceReports/${report.id}/attendanceRecords`
            )
            .get();
          detailed.push({
            reportId: report.id,
            records: records.value,
          });
        }

        return {
          content: [{ type: "text" as const, text: JSON.stringify(detailed, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            {
              type: "text" as const,
              text: `Failed to get attendance: ${message}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "get_meeting_transcripts",
    "Get transcripts for a meeting (requires Teams Premium recording)",
    GetMeetingTranscriptsSchema.shape,
    async ({ meetingId }) => {
      try {
        const client = await graphService.getClient();
        const result = await client
          .api(`/me/onlineMeetings/${meetingId}/transcripts`)
          .get();

        const transcripts = result.value as Array<{
          id: string;
          createdDateTime: string;
        }>;
        const detailed = [];

        for (const transcript of transcripts) {
          try {
            const content = await client
              .api(
                `/me/onlineMeetings/${meetingId}/transcripts/${transcript.id}/content`
              )
              .get();
            detailed.push({
              transcriptId: transcript.id,
              createdDateTime: transcript.createdDateTime,
              content: JSON.stringify(content),
            });
          } catch {
            detailed.push({
              transcriptId: transcript.id,
              createdDateTime: transcript.createdDateTime,
              content: "Unable to fetch transcript content",
            });
          }
        }

        return {
          content: [{ type: "text" as const, text: JSON.stringify(detailed, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            {
              type: "text" as const,
              text: `Failed to get transcripts: ${message}`,
            },
          ],
          isError: true,
        };
      }
    }
  );
}
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors.

---

### Task 12: Messages Tools

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\tools\messages.ts`

- [ ] **Step 1: Write tools/messages.ts**

```typescript
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import { SearchMessagesSchema } from "../types.js";

export function registerMessagesTools(
  server: McpServer,
  graphService: GraphService
): void {
  server.tool(
    "search_messages",
    "Search across all Teams messages using Microsoft Search API",
    SearchMessagesSchema.shape,
    async ({ query, limit }) => {
      try {
        const client = await graphService.getClient();
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
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors.

---

### Task 13: Entry Point (CLI)

**Files:**
- Create: `D:\2026\mcp-teams-from-scratch\src\index.ts`

- [ ] **Step 1: Write index.ts**

```typescript
#!/usr/bin/env node

import "dotenv/config";
import { startHttpServer, startStdioServer } from "./server.js";

async function main() {
  const args = process.argv.slice(2);
  const useHttp = args.includes("--http") || process.env.TEAMS_MCP_TRANSPORT === "http";
  const useOAuth = args.includes("--oauth") || process.env.TEAMS_MCP_OAUTH === "true";

  if (args.includes("--help") || args.includes("-h") || args.includes("help")) {
    console.log("MCP Teams Server");
    console.log("");
    console.log("Usage:");
    console.log("  npx tsx src/index.ts               Start MCP server (stdio)");
    console.log("  npx tsx src/index.ts --http        Start MCP server (HTTP)");
    console.log("  npx tsx src/index.ts --http --oauth  HTTP mode with OAuth 2.1");
    console.log("");
    console.log("Environment variables:");
    console.log("  AZURE_CLIENT_ID       Entra App client ID");
    console.log("  AZURE_CLIENT_SECRET   Entra App client secret");
    console.log("  AZURE_TENANT_ID       Entra tenant ID");
    console.log("  TEAMS_MCP_PORT=3000   HTTP server port");
    console.log("  TEAMS_MCP_TRANSPORT=http  Use HTTP transport");
    console.log("  TEAMS_MCP_OAUTH=true     Enable OAuth 2.1 (HTTP only)");
    return;
  }

  if (useHttp) {
    await startHttpServer(useOAuth);
  } else {
    await startStdioServer();
  }
}

main().catch((error) => {
  console.error("Failed to start:", error);
  process.exit(1);
});
```

- [ ] **Step 2: Run full TypeScript build check**

```bash
cd D:\2026\mcp-teams-from-scratch; npx tsc --noEmit
```
Expected: No errors across all files.

- [ ] **Step 3: Run full build**

```bash
cd D:\2026\mcp-teams-from-scratch; npm run build
```
Expected: Successful compilation, `dist/` directory created with all JS outputs.

---
