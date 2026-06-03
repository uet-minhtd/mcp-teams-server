# MCP Teams Server

MCP server tích hợp Microsoft Teams qua Microsoft Graph API. Cho phép AI assistant (Claude Desktop, VS Code, Cursor) tương tác với Teams: đọc/gửi tin nhắn, quản lý kênh, chat, meeting, tìm kiếm người dùng.

Viết từ đầu với TypeScript, không phụ thuộc thư viện cá nhân bên ngoài.

## Kiến trúc

```text
MCP Client (Claude Desktop, VS Code, Cursor)
        │
        │  MCP OAuth 2.1 (Bearer token, xác thực qua Microsoft)
        ▼
  MCP Teams Server ──── Express + @modelcontextprotocol/sdk
        │
        │  ClientSecretCredential (Entra App registration)
        ▼
  Microsoft Graph API ──── @microsoft/microsoft-graph-client
        │
        ▼
  Microsoft Teams (users, teams, channels, chats, meetings, messages)
```

## Yêu cầu

- Node.js 20+
- Entra App Registration với các quyền Microsoft Graph (xem bảng Tools bên dưới)
- Microsoft 365 account có Teams

## Cài đặt

```bash
git clone <repo-url>
cd mcp-teams-from-scratch
npm install
npm run build
```

## Cấu hình

Copy `.env.example` thành `.env` và điền thông tin Entra App:

```env
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
TEAMS_MCP_PORT=8888
TEAMS_MCP_OAUTH=true
```

Lấy các giá trị này từ **Azure Portal** → **Microsoft Entra ID** → **App registrations** → app của bạn.

### Quyền Microsoft Graph cần cấp (API Permissions)

| Permission | Type | Dùng cho |
|------------|------|----------|
| User.Read | Delegated | auth_status, get_current_user |
| User.ReadBasic.All | Application | search_users |
| Team.ReadBasic.All | Application | list_teams |
| Channel.ReadBasic.All | Application | list_channels |
| ChannelMessage.Read.All | Application | get_channel_messages, search_messages |
| ChannelMessage.Send | Application | send_channel_message |
| Chat.Read | Application | list_chats, get_chat_messages, search_messages |
| Chat.ReadWrite | Application | send_chat_message |
| OnlineMeetings.Read | Application | list_meetings, get_meeting |
| OnlineMeetingArtifact.Read.All | Application | get_meeting_attendance, get_meeting_transcripts |

Sau khi thêm permissions, nhấn **Grant admin consent**.

## Chạy

### Development (tsx hot-reload)

```bash
# STDIO mode (cho local MCP client như Claude Desktop)
npm run dev

# HTTP mode (không OAuth — dùng cho test nội bộ)
npm run dev -- --http

# HTTP mode + OAuth 2.1 (cho remote client như VS Code, Cursor)
npm run dev -- --http --oauth
```

### Production

```bash
npm run build
npm start
```

### Docker

```bash
docker compose up -d      # build + start
docker compose logs -f    # xem logs
docker compose down       # dừng
```

MCP endpoint: `http://localhost:3000/mcp`

## Cấu hình MCP Client

### Claude Desktop / Claude Code

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "npx",
      "args": ["tsx", "D:\\2026\\mcp-teams-from-scratch\\src\\index.ts"],
      "env": {
        "AZURE_CLIENT_ID": "xxx",
        "AZURE_CLIENT_SECRET": "xxx",
        "AZURE_TENANT_ID": "xxx"
      }
    }
  }
}
```

### HTTP mode (VS Code, Cursor, remote clients)

```json
{
  "mcpServers": {
    "teams-mcp": {
      "url": "http://localhost:3000/mcp"
    }
  }
}
```

## MCP Tools (16)

### Auth
| Tool | Mô tả |
|------|-------|
| `auth_status` | Kiểm tra trạng thái xác thực với Microsoft Graph |

### Users
| Tool | Mô tả |
|------|-------|
| `get_current_user` | Lấy thông tin user đang đăng nhập |
| `search_users` | Tìm kiếm user theo tên hoặc email |

### Teams & Channels
| Tool | Mô tả |
|------|-------|
| `list_teams` | Liệt kê tất cả teams user là thành viên |
| `list_channels` | Liệt kê channels trong một team |
| `get_channel_messages` | Đọc tin nhắn từ channel |
| `send_channel_message` | Gửi tin nhắn vào channel |

### Chats
| Tool | Mô tả |
|------|-------|
| `list_chats` | Liệt kê tất cả chats (1:1 và nhóm) |
| `get_chat_messages` | Đọc tin nhắn từ chat |
| `send_chat_message` | Gửi tin nhắn vào chat |

### Meetings
| Tool | Mô tả |
|------|-------|
| `list_meetings` | Liệt kê online meetings |
| `get_meeting` | Chi tiết meeting (link join, subject, organizer) |
| `get_meeting_attendance` | Báo cáo điểm danh + danh sách tham gia |
| `get_meeting_transcripts` | Transcript nội dung meeting (cần Teams Premium) |

### Search
| Tool | Mô tả |
|------|-------|
| `search_messages` | Tìm kiếm toàn bộ tin nhắn Teams |

## Luồng xác thực

### MCP OAuth 2.1 (Client → Server)

```
1. Client gửi request không token → 401 + WWW-Authenticate
2. Client lấy Protected Resource Metadata (RFC 9728)
3. Client lấy Authorization Server Metadata (RFC 8414)
4. Client mở browser → Microsoft login → redirect callback → authorization code
5. Client đổi code lấy access token tại /token
6. Client gửi Authorization: Bearer <token> trong mọi MCP request
7. Server validate token bằng cách gọi GET https://graph.microsoft.com/v1.0/me
```

### Graph Auth (Server → Microsoft Graph)

Server dùng `ClientSecretCredential` (`AZURE_CLIENT_ID` + `AZURE_CLIENT_SECRET` + `AZURE_TENANT_ID`) để lấy access token gọi Graph API. Token được cache trong memory và tự động refresh.

## Cấu trúc thư mục

```
src/
├── index.ts              # Entry point: CLI (--http, --oauth, --help)
├── server.ts             # McpServer factory + STDIO/HTTP transport
├── types.ts              # Zod schemas cho tool inputs
├── auth/
│   ├── middleware.ts     # requireBearerAuth middleware
│   └── endpoints.ts     # OAuth 2.1 endpoints
├── graph/
│   └── client.ts         # GraphService singleton (ClientSecretCredential)
└── tools/
    ├── auth.ts           # auth_status
    ├── users.ts          # get_current_user, search_users
    ├── teams.ts          # list_teams, list_channels, get/send_channel_message
    ├── chats.ts          # list_chats, get/send_chat_message
    ├── meetings.ts       # list_meetings, get_meeting, attendance, transcripts
    └── messages.ts       # search_messages
```

## Tech Stack

| Package | Purpose |
|---------|---------|
| `@modelcontextprotocol/sdk` | MCP server, streamable HTTP transport |
| `@microsoft/microsoft-graph-client` | Microsoft Graph API client |
| `@azure/identity` | ClientSecretCredential cho Graph auth |
| `express` | HTTP server |
| `zod` | Input validation cho MCP tools |
| `dotenv` | Environment variables |
| `typescript` + `tsx` | Build + dev |

## License

MIT
