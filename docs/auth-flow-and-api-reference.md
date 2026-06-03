# Auth Flow & API Reference

## Luồng Authentication (OAuth 2.1 Flow)

```
MCP Client (AI Agent / Claude Desktop)
         │
         │  1. Gọi POST /mcp (không có token)
         ▼
┌─────────────────────────────────────────────────┐
│              Express HTTP Server                │
│                                                 │
│  requireBearerAuth() middleware                 │
│  → Phát hiện thiếu token                        │
│  → Trả về 401 + WWW-Authenticate header         │
│    (chứa URL của OAuth server)                  │
└─────────────────────┬───────────────────────────┘
                      │
                      │  2. Client đọc WWW-Authenticate header
                      │     → Gọi GET /.well-known/oauth-authorization-server
                      ▼
┌─────────────────────────────────────────────────┐
│           OAuth Discovery Endpoints             │
│  Trả về: authorization_endpoint, token_endpoint │
│          registration_endpoint, scopes...        │
└─────────────────────┬───────────────────────────┘
                      │
                      │  3. POST /register (Dynamic Client Registration)
                      ▼
┌─────────────────────────────────────────────────┐
│  dcrRegisterHandler                             │
│  → Trả về client_id (AZURE_CLIENT_ID)           │
└─────────────────────┬───────────────────────────┘
                      │
                      │  4. GET /authorize?client_id=...&redirect_uri=...&state=...
                      ▼
┌─────────────────────────────────────────────────┐
│  authorizeHandler                               │
│  → Lưu state → pendingCallbacks Map             │
│  → Thay redirect_uri = localhost:{port}/callback│
│  → Redirect 302 sang Microsoft OAuth            │
└─────────────────────┬───────────────────────────┘
                      │
                      │  5. Browser user đăng nhập Microsoft
                      ▼
┌─────────────────────────────────────────────────┐
│         Microsoft Identity Platform             │
│  https://login.microsoftonline.com/common       │
│  /oauth2/v2.0/authorize                         │
└─────────────────────┬───────────────────────────┘
                      │
                      │  6. Microsoft callback → GET /callback?code=...&state=...
                      ▼
┌─────────────────────────────────────────────────┐
│  oauthCallbackHandler                           │
│  → Tra state trong pendingCallbacks             │
│  → Redirect về redirect_uri gốc của client     │
│    kèm authorization code                       │
└─────────────────────┬───────────────────────────┘
                      │
                      │  7. Client POST /token với authorization code
                      ▼
┌─────────────────────────────────────────────────┐
│  tokenExchangeHandler                           │
│  → Forward request đến Microsoft token endpoint │
│  → Trả về access_token + refresh_token          │
└─────────────────────┬───────────────────────────┘
                      │
                      │  8. Client gọi POST /mcp với Bearer token
                      ▼
┌─────────────────────────────────────────────────┐
│  requireBearerAuth() middleware                 │
│  → Gọi GraphService.validateToken(token)        │
│    ├─ Decode JWT, check exp                     │
│    ├─ Kiểm tra cache (TTL 5 phút)               │
│    └─ Nếu miss cache → GET /me từ Graph API     │
│  → Attach AuthContext vào req.auth              │
│  → next() → MCP Handler                        │
└─────────────────────┬───────────────────────────┘
                      │
                      │  9. MCP Server xử lý tool calls
                      ▼
┌─────────────────────────────────────────────────┐
│  McpServer (per-session)                        │
│  → Dùng userToken để tạo Graph client           │
│  → Gọi Microsoft Graph API với delegated perms  │
└─────────────────────────────────────────────────┘
```

---

## OAuth 2.1 Endpoints (HTTP Server)

### `GET /.well-known/oauth-protected-resource`

**Mục đích:** Thông báo với MCP client rằng server yêu cầu OAuth, trỏ về địa chỉ authorization server.

**Cách dùng:** MCP client tự động gọi khi nhận phản hồi 401 từ `/mcp`.

**Response:**
```json
{
  "authorization_servers": ["http://localhost:3000"]
}
```

---

### `GET /.well-known/oauth-authorization-server`

**Mục đích:** OAuth discovery metadata — cung cấp toàn bộ cấu hình OAuth server (endpoints, scopes, grant types).

**Cách dùng:** MCP client tự động gọi sau khi đọc `oauth-protected-resource`.

**Response:**
```json
{
  "issuer": "http://localhost:3000",
  "authorization_endpoint": "http://localhost:3000/authorize",
  "token_endpoint": "http://localhost:3000/token",
  "registration_endpoint": "http://localhost:3000/register",
  "response_types_supported": ["code"],
  "code_challenge_methods_supported": ["S256"],
  "grant_types_supported": ["authorization_code", "refresh_token"],
  "scopes_supported": ["openid", "profile", "offline_access", "User.Read", "..."]
}
```

---

### `POST /register`

**Mục đích:** Dynamic Client Registration (DCR) — client đăng ký để nhận `client_id` trước khi bắt đầu OAuth flow.

**Cách dùng:** Client POST không cần body. Server trả về `client_id` là `AZURE_CLIENT_ID` đã cấu hình.

**Response:**
```json
{
  "client_id": "14d82eec-204b-4c2f-b7e8-296a70dab67e",
  "client_id_issued_at": 1748000000,
  "client_secret_expires_at": 0,
  "grant_types": ["authorization_code", "refresh_token"],
  "response_types": ["code"],
  "redirect_uris": ["http://localhost", "http://127.0.0.1"],
  "scope": "openid profile offline_access User.Read ..."
}
```

---

### `GET /authorize`

**Mục đích:** Khởi động OAuth authorization flow. Server proxy request sang Microsoft Identity Platform, đồng thời lưu `state` để relay callback.

**Cách dùng:** Browser hoặc MCP client mở URL này với các query params chuẩn OAuth.

**Query params:**
| Param | Bắt buộc | Mô tả |
|-------|----------|-------|
| `client_id` | Có | Client ID lấy từ `/register` |
| `redirect_uri` | Có | URI để nhận authorization code sau khi đăng nhập |
| `state` | Có | Random string chống CSRF |
| `code_challenge` | Có | PKCE challenge (S256) |
| `response_type` | Có | Luôn là `code` |
| `scope` | Không | Danh sách scope, mặc định là toàn bộ DCR scopes |

**Lưu ý:** Server tự thay `redirect_uri` thành `localhost:{port}/callback` trước khi forward sang Microsoft. State được lưu tối đa **5 phút**, giới hạn **100 pending** đồng thời.

---

### `GET /callback`

**Mục đích:** Nhận authorization code từ Microsoft sau khi user đăng nhập, sau đó relay về `redirect_uri` gốc của client.

**Cách dùng:** Microsoft tự động gọi endpoint này (không phải client gọi trực tiếp).

**Query params nhận từ Microsoft:**
| Param | Mô tả |
|-------|-------|
| `code` | Authorization code (dùng để đổi token) |
| `state` | State gốc để tra cứu `redirect_uri` |
| `error` | Nếu có lỗi (vd: `access_denied`) |

---

### `POST /token`

**Mục đích:** Đổi authorization code lấy `access_token` và `refresh_token`. Cũng dùng để refresh token khi hết hạn.

**Cách dùng:** Client POST với `Content-Type: application/x-www-form-urlencoded`.

**Body (authorization_code):**
```
grant_type=authorization_code
&code=<authorization_code>
&client_id=<client_id>
&code_verifier=<pkce_verifier>
```

**Body (refresh_token):**
```
grant_type=refresh_token
&refresh_token=<refresh_token>
&client_id=<client_id>
```

**Response (từ Microsoft, được proxy lại):**
```json
{
  "access_token": "eyJ...",
  "refresh_token": "0.A...",
  "token_type": "Bearer",
  "expires_in": 3600,
  "scope": "openid profile User.Read ..."
}
```

---

## MCP Transport Endpoints

### `POST /mcp`

**Mục đích:** Gửi MCP JSON-RPC request đến server (initialize session, gọi tool...).

**Cách dùng:**
```
POST /mcp
Authorization: Bearer <access_token>
Content-Type: application/json
mcp-session-id: <session_id>  (trừ lần initialize đầu tiên)

{ "jsonrpc": "2.0", "method": "tools/call", "params": { ... }, "id": 1 }
```

- Lần đầu tiên (initialize): không cần `mcp-session-id`, server sẽ trả về session ID mới.
- Các lần sau: phải gửi kèm `mcp-session-id` nhận từ lần initialize.

---

### `GET /mcp`

**Mục đích:** Mở SSE (Server-Sent Events) stream để nhận response bất đồng bộ từ server.

**Cách dùng:**
```
GET /mcp
Authorization: Bearer <access_token>
mcp-session-id: <session_id>
Accept: text/event-stream
```

---

### `DELETE /mcp`

**Mục đích:** Đóng và xoá session hiện tại, giải phóng tài nguyên server.

**Cách dùng:**
```
DELETE /mcp
Authorization: Bearer <access_token>
mcp-session-id: <session_id>
```

---

## MCP Tools (Microsoft Graph API)

### 🔐 Auth

#### `auth_status`
Kiểm tra token hiện tại có hợp lệ không và lấy thông tin user đang đăng nhập.

**Input:** Không có

**Output:**
```json
{
  "authenticated": true,
  "userPrincipalName": "user@company.com",
  "displayName": "Nguyen Van A"
}
```

---

### 👤 Users

#### `get_current_user`
Lấy profile đầy đủ của user đang đăng nhập từ Microsoft Graph `/me`.

**Input:** Không có

**Output:** Toàn bộ user object từ Graph API (displayName, mail, jobTitle, department, v.v.)

---

#### `search_users`
Tìm kiếm user trong tổ chức theo tên hoặc email.

**Input:**
| Param | Kiểu | Mô tả |
|-------|------|-------|
| `query` | string | Tên hoặc email cần tìm (prefix match) |

**Ví dụ:** `{ "query": "nguyen" }` → tìm tất cả user có displayName hoặc email bắt đầu bằng "nguyen"

---

### 👥 Teams & Channels

#### `list_teams`
Liệt kê tất cả Microsoft Teams mà user hiện tại đang là thành viên.

**Input:** Không có

**Output:** Danh sách teams với id, displayName, description

---

#### `list_channels`
Liệt kê tất cả channels trong một team cụ thể.

**Input:**
| Param | Kiểu | Mô tả |
|-------|------|-------|
| `teamId` | string | ID của team (lấy từ `list_teams`) |

---

#### `get_channel_messages`
Lấy các tin nhắn gần nhất từ một channel.

**Input:**
| Param | Kiểu | Mặc định | Mô tả |
|-------|------|----------|-------|
| `teamId` | string | - | ID của team |
| `channelId` | string | - | ID của channel (lấy từ `list_channels`) |
| `limit` | number | 20 | Số lượng tin nhắn tối đa |

---

#### `send_channel_message`
Gửi một tin nhắn văn bản vào channel.

**Input:**
| Param | Kiểu | Mô tả |
|-------|------|-------|
| `teamId` | string | ID của team |
| `channelId` | string | ID của channel |
| `message` | string | Nội dung tin nhắn |

**Output:** `{ "success": true, "messageId": "..." }`

---

### 💬 Chats

#### `list_chats`
Liệt kê tất cả cuộc trò chuyện (1-1 và group chat) mà user đang tham gia.

**Input:** Không có

**Output:** Danh sách chat với id, chatType, topic

---

#### `get_chat_messages`
Lấy các tin nhắn gần nhất từ một chat.

**Input:**
| Param | Kiểu | Mặc định | Mô tả |
|-------|------|----------|-------|
| `chatId` | string | - | ID của chat (lấy từ `list_chats`) |
| `limit` | number | 20 | Số lượng tin nhắn tối đa |

---

#### `send_chat_message`
Gửi tin nhắn vào một chat.

**Input:**
| Param | Kiểu | Mô tả |
|-------|------|-------|
| `chatId` | string | ID của chat |
| `message` | string | Nội dung tin nhắn |

**Output:** `{ "success": true, "messageId": "..." }`

---

### 📅 Meetings

#### `list_meetings`
Liệt kê các online meeting (cả đã qua và sắp tới) của user.

**Input:**
| Param | Kiểu | Mặc định | Mô tả |
|-------|------|----------|-------|
| `limit` | number | 20 | Số lượng meeting tối đa |

---

#### `get_meeting`
Lấy thông tin chi tiết của một meeting cụ thể.

**Input:**
| Param | Kiểu | Mô tả |
|-------|------|-------|
| `meetingId` | string | ID của meeting (lấy từ `list_meetings`) |

**Output:** Meeting object với subject, startDateTime, endDateTime, joinUrl, participants, v.v.

---

#### `get_meeting_attendance`
Lấy báo cáo điểm danh đầy đủ của meeting, bao gồm danh sách người tham dự và thời gian tham gia.

**Input:**
| Param | Kiểu | Mô tả |
|-------|------|-------|
| `meetingId` | string | ID của meeting |

**Output:** Danh sách attendance reports, mỗi report có attendanceRecords với thông tin từng người.

---

#### `get_meeting_transcripts`
Lấy nội dung transcript của meeting (yêu cầu **Teams Premium** và tính năng recording đã bật).

**Input:**
| Param | Kiểu | Mô tả |
|-------|------|-------|
| `meetingId` | string | ID của meeting |

**Output:** Danh sách transcripts với nội dung và thời gian tạo.

---

### 🔍 Search

#### `search_messages`
Tìm kiếm tin nhắn trên toàn bộ Teams (channels + chats) sử dụng Microsoft Search API với cú pháp KQL.

**Input:**
| Param | Kiểu | Mặc định | Mô tả |
|-------|------|----------|-------|
| `query` | string | - | Từ khóa tìm kiếm, hỗ trợ KQL syntax |
| `limit` | number | 20 | Số kết quả tối đa |

**Ví dụ KQL:**
- `"budget report"` — tìm cụm từ chính xác
- `from:user@company.com` — tìm tin nhắn từ người cụ thể
- `"quarterly review" AND from:manager@company.com` — kết hợp điều kiện

---

## Ghi chú kỹ thuật

### Token Validation Flow

`GraphService.validateToken(token)` thực hiện theo thứ tự:

1. **Decode JWT** — kiểm tra `exp` claim, từ chối ngay nếu đã hết hạn
2. **Opaque token** — nếu không decode được (personal account), bỏ qua bước 1
3. **Cache lookup** — nếu token đã được validate trong **5 phút** qua, trả về kết quả cache
4. **Graph API call** — `GET https://graph.microsoft.com/v1.0/me` với Bearer token
5. **Degraded mode** — nếu Graph call fail nhưng JWT còn hạn, vẫn cho phép truy cập dựa trên thông tin trong payload JWT (cache 60 giây)

### Scopes yêu cầu

| Scope | Dùng cho tool |
|-------|--------------|
| `User.Read` | `get_current_user`, `auth_status` |
| `User.ReadBasic.All` | `search_users` |
| `Team.ReadBasic.All` | `list_teams` |
| `Channel.ReadBasic.All` | `list_channels` |
| `ChannelMessage.Read.All` | `get_channel_messages` |
| `ChannelMessage.Send` | `send_channel_message` |
| `Chat.Read` | `list_chats`, `get_chat_messages` |
| `Chat.ReadWrite` | `send_chat_message` |
| `OnlineMeetings.Read` | `list_meetings`, `get_meeting` |
| `OnlineMeetingArtifact.Read.All` | `get_meeting_attendance`, `get_meeting_transcripts` |

### Chế độ vận hành

| Mode | Lệnh khởi động | Auth | Dùng khi |
|------|---------------|------|----------|
| stdio | `npx tsx src/index.ts` | Không (app token) | Local, Claude Desktop |
| HTTP | `--http` | Không | Remote không cần login |
| HTTP + OAuth | `--http --oauth` | OAuth 2.1 | Remote, multi-user |
