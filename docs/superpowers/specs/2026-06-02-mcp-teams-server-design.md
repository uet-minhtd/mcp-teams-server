# MCP Teams Server — Design Spec

**Date:** 2026-06-02
**Status:** Draft

## Overview

A custom MCP server that provides Microsoft Teams integration via Microsoft Graph API. Built from scratch for learning purposes — no dependency on personal/external libraries beyond official SDKs.

## Architecture

```text
MCP Client (Claude Desktop, VS Code, etc.)
        │
        │  MCP OAuth 2.1 (Bearer token via Microsoft login)
        ▼
  MCP Teams Server  ──── Express + @modelcontextprotocol/sdk
        │
        │  Client Secret Credential (Entra App registration)
        ▼
  Microsoft Graph API ──── @microsoft/microsoft-graph-client
        │
        ▼
  Microsoft Teams data (users, teams, channels, chats, meetings, messages)
```

## Two Authentication Layers

### 1. MCP Auth (Client → Server)

The MCP server acts as its own OAuth Authorization Server, delegating actual user authentication to Microsoft:

1. Client sends unauthenticated request → 401 with `WWW-Authenticate` header
2. Client fetches `/.well-known/oauth-protected-resource` → discovers AS
3. Client fetches `/.well-known/oauth-authorization-server` → endpoints
4. Client opens browser → Microsoft login → redirect callback → authorization code
5. Client exchanges code for token at `/token`
6. Client includes `Authorization: Bearer <token>` in all MCP requests
7. Server validates token by calling `GET https://graph.microsoft.com/v1.0/me`

### 2. Graph Auth (Server → Microsoft Graph)

Server uses `@azure/identity` `ClientSecretCredential` with values from `.env`:
- `AZURE_CLIENT_ID` — Entra App Application ID
- `AZURE_CLIENT_SECRET` — Entra App client secret
- `AZURE_TENANT_ID` — Entra tenant ID

Token is cached in memory and refreshed automatically.

## Project Structure

```
mcp-teams-from-scratch/
├── .env
├── .env.example
├── package.json
├── tsconfig.json
├── src/
│   ├── index.ts              # Entry point: CLI parsing, startup
│   ├── server.ts             # MCP server + HTTP transport + Express
│   ├── auth/
│   │   ├── middleware.ts     # requireBearerAuth — validates token via Graph /me
│   │   └── endpoints.ts     # OAuth 2.1: metadata, authorize, callback, token, dcr
│   ├── graph/
│   │   └── client.ts         # GraphService singleton — ClientSecretCredential + Graph Client
│   ├── tools/
│   │   ├── auth.ts           # auth_status
│   │   ├── users.ts          # get_current_user, search_users
│   │   ├── teams.ts          # list_teams, list_channels, get_channel_messages, send_channel_message
│   │   ├── chats.ts          # list_chats, get_chat_messages, send_chat_message
│   │   ├── meetings.ts       # list_meetings, get_meeting, get_meeting_attendance, get_meeting_transcripts
│   │   └── messages.ts       # search_messages
│   └── types.ts              # Shared types (tool inputs, Graph response types)
```

## MCP Tools

| Category  | Tool                      | Description                                    | Graph Permission                  |
|-----------|---------------------------|------------------------------------------------|----------------------------------|
| Auth      | `auth_status`             | Check authentication status                    | User.Read                        |
| Users     | `get_current_user`        | Get authenticated user info                    | User.Read                        |
| Users     | `search_users`            | Search users by name/email                     | User.ReadBasic.All               |
| Teams     | `list_teams`              | List joined teams                              | Team.ReadBasic.All               |
| Teams     | `list_channels`           | List channels in a team                        | Channel.ReadBasic.All            |
| Teams     | `get_channel_messages`    | Get messages from a channel                    | ChannelMessage.Read.All          |
| Teams     | `send_channel_message`    | Send message to a channel                      | ChannelMessage.Send              |
| Chats     | `list_chats`              | List user's chats                              | Chat.Read                        |
| Chats     | `get_chat_messages`       | Get messages from a chat                       | Chat.Read                        |
| Chats     | `send_chat_message`       | Send message to a chat                         | Chat.ReadWrite                   |
| Meetings  | `list_meetings`           | List online meetings (upcoming/past)           | OnlineMeetings.Read              |
| Meetings  | `get_meeting`             | Get meeting details (URL, subject, organizer)  | OnlineMeetings.Read              |
| Meetings  | `get_meeting_attendance`  | Get attendance reports + records               | OnlineMeetingArtifact.Read.All   |
| Meetings  | `get_meeting_transcripts` | Get meeting transcripts                        | OnlineMeetingArtifact.Read.All   |
| Messages  | `search_messages`         | Search across Teams messages                   | ChannelMessage.Read.All, Chat.Read|

## Data Flow: Tool Execution

```text
1. LLM decides to call tool → MCP Client sends tools/call JSON-RPC
2. Express POST /mcp → requireBearerAuth middleware validates token
3. Server routes to tool handler
4. Handler calls GraphService.getClient()
5. GraphService uses ClientSecretCredential to get access token
6. Graph Client calls Microsoft Graph API
7. Result formatted and returned as MCP content
```

## Dependencies

| Package                           | Purpose                              |
|-----------------------------------|--------------------------------------|
| `@modelcontextprotocol/sdk`       | MCP server, HTTP transport, types    |
| `@microsoft/microsoft-graph-client` | Microsoft Graph API client          |
| `@azure/identity`                 | ClientSecretCredential for Graph auth |
| `express`                         | HTTP server                          |
| `dotenv`                          | Environment variable loading         |
| `typescript`, `tsx`               | Build + dev runner                   |

## Error Handling

- Graph API errors → mapped to MCP error responses with meaningful messages
- Auth failures → 401 with proper WWW-Authenticate header
- Token refresh failures → clear state, return auth error
- Unknown errors → 500 with generic message, log details server-side

## Scope & Boundaries

**In scope:**
- STDIO transport (for local use) + HTTP transport with OAuth (for remote use)
- All 16 tools listed above
- Single-tenant operation (one Entra App)
- In-memory state (no persistence beyond env vars)

**Out of scope (for now):**
- File upload/download
- Message edit/delete
- Multi-tenant support
- Database/persistent storage
- Tests (added later)
