# Hướng dẫn triển khai MCP Teams Server cho doanh nghiệp

## Tổng quan

MCP Teams Server kết nối AI agents (qua Model Context Protocol) với Microsoft Teams thông qua Microsoft Graph API. Hướng dẫn này mô tả các bước cấu hình để triển khai trong môi trường doanh nghiệp.

## Yêu cầu

- Tài khoản Microsoft 365 work/school (có Teams license)
- Quyền Global Admin hoặc Privileged Role Admin trên Azure Entra ID
- Node.js >= 20
- HTTPS domain cho production (hoặc reverse proxy)

---

## 1. Azure App Registration

### a) Tạo App Registration

1. Truy cập **Azure Portal → Microsoft Entra ID → App registrations → New registration**
2. Cấu hình:
   - **Name**: MCP Teams Server (hoặc tên tùy chọn)
   - **Supported account types**: Accounts in this organizational directory only (Single tenant)
   - **Redirect URI**: Web — `https://your-domain.com/callback`
3. Bấm **Register**

### b) API Permissions

Vào **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, thêm:

| Permission | Mô tả | Admin Consent |
|---|---|---|
| `User.Read` | Đọc profile user đăng nhập | Không |
| `User.ReadBasic.All` | Tìm kiếm users trong tổ chức | Có |
| `Team.ReadBasic.All` | Liệt kê teams user tham gia | Có |
| `Channel.ReadBasic.All` | Liệt kê channels trong team | Có |
| `ChannelMessage.Read.All` | Đọc tin nhắn channel | Có |
| `ChannelMessage.Send` | Gửi tin nhắn vào channel | Có |
| `Chat.Read` | Đọc chats của user | Có |
| `Chat.ReadWrite` | Gửi tin nhắn chat | Có |
| `TeamMember.Read.All` | Đọc thành viên team | Có |
| `OnlineMeetings.Read` | Đọc online meetings | Có |
| `OnlineMeetingArtifact.Read.All` | Đọc transcript và attendance | Có |

### c) Grant Admin Consent

Bấm **"Grant admin consent for [Org Name]"** — cần quyền Global Admin hoặc Privileged Role Admin.

### d) Client Secret

1. Vào **Certificates & secrets → Client secrets → New client secret**
2. Đặt mô tả và thời hạn (khuyến nghị 12-24 tháng)
3. **Copy ngay giá trị Value** — chỉ hiện 1 lần

> **Production**: Khuyến nghị dùng Certificate thay vì Client Secret, và lưu trong Azure Key Vault.

---

## 2. Cấu hình Environment Variables

### File `.env` cho production:

```env
# Azure App Registration
AZURE_CLIENT_ID=<application-client-id>
AZURE_CLIENT_SECRET=<client-secret-value>
AZURE_TENANT_ID=<organization-tenant-id>

# Server
TEAMS_MCP_PORT=8888
TEAMS_MCP_OAUTH=true
NODE_ENV=production
```

### Lấy giá trị:

| Biến | Vị trí trên Azure Portal |
|---|---|
| `AZURE_CLIENT_ID` | App Registration → Overview → Application (client) ID |
| `AZURE_CLIENT_SECRET` | App Registration → Certificates & secrets → Value |
| `AZURE_TENANT_ID` | App Registration → Overview → Directory (tenant) ID |

---

## 3. Redirect URI

### Trên Azure Portal (Authentication → Redirect URIs):

| Môi trường | Redirect URI |
|---|---|
| Development | `http://localhost:8888/callback` |
| Production | `https://mcp-teams.company.com/callback` |

### Trong code:

Cập nhật `src/auth/endpoints.ts` để dùng dynamic host cho production:

```typescript
// Thay vì hardcode localhost:
const proxyRu = `http://localhost:${port}/callback`;

// Dùng dynamic host:
const proxyRu = `${req.protocol}://${req.get("host")}/callback`;
```

---

## 4. Tenant-specific Authority

Trong `src/auth/endpoints.ts`, đổi authority sang tenant cụ thể:

```typescript
// Development (multi-tenant, hỗ trợ personal + work):
const MICROSOFT_AUTHORITY = "https://login.microsoftonline.com/common";

// Production (single tenant, chỉ tổ chức):
const MICROSOFT_AUTHORITY = `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`;
```

---

## 5. Deployment

### Docker

```bash
# Build
docker build -t mcp-teams-server .

# Run
docker run -d \
  --name mcp-teams \
  -p 8888:8888 \
  --env-file .env.production \
  --restart unless-stopped \
  mcp-teams-server
```

### Docker Compose

```yaml
services:
  mcp-teams:
    build: .
    ports:
      - "8888:8888"
    env_file:
      - .env.production
    restart: unless-stopped
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:8888/.well-known/oauth-protected-resource"]
      interval: 30s
      timeout: 10s
      retries: 3
```

### Azure App Service

1. Tạo Web App (Linux, Node 20)
2. Set environment variables trong Configuration
3. Deploy qua GitHub Actions hoặc Azure CLI

---

## 6. Reverse Proxy (HTTPS)

### Nginx

```nginx
server {
    listen 443 ssl;
    server_name mcp-teams.company.com;

    ssl_certificate /etc/ssl/certs/your-cert.pem;
    ssl_certificate_key /etc/ssl/private/your-key.pem;

    location / {
        proxy_pass http://localhost:8888;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

### Caddy (tự động HTTPS)

```
mcp-teams.company.com {
    reverse_proxy localhost:8888
}
```

---

## 7. Security Checklist

| Mục | Cấu hình |
|-----|----------|
| HTTPS | Bắt buộc cho production |
| Tenant isolation | Dùng single-tenant app + tenant-specific authority |
| Secret management | Lưu secrets trong Key Vault, không commit vào git |
| Rate limiting | Thêm middleware chống abuse (express-rate-limit) |
| Token cache | Đã built-in (5 phút) — dùng Redis cho multi-instance |
| CORS | Restrict allowed origins nếu expose public |
| Logging | Redirect stdout → log aggregator (CloudWatch, ELK, Datadog) |
| Monitoring | Health check endpoint + uptime monitoring |
| Network | Restrict inbound traffic bằng NSG/firewall |

---

## 8. Kiro MCP Client Config

### Cho developer trong team (file `~/.kiro/settings/mcp.json` hoặc workspace `.kiro/settings/mcp.json`):

```json
{
  "mcpServers": {
    "mcp-teams": {
      "url": "https://mcp-teams.company.com/mcp",
      "disabled": false,
      "autoApprove": ["auth_status", "get_current_user", "list_teams", "list_channels"]
    }
  }
}
```

---

## 9. Available MCP Tools

| Tool | Mô tả | Params |
|------|--------|--------|
| `auth_status` | Kiểm tra trạng thái xác thực | — |
| `get_current_user` | Lấy profile người dùng | — |
| `search_users` | Tìm kiếm user theo tên/email | `query` |
| `list_teams` | Liệt kê teams user tham gia | — |
| `list_channels` | Liệt kê channels trong team | `teamId` |
| `get_channel_messages` | Đọc tin nhắn channel | `teamId`, `channelId`, `limit` |
| `send_channel_message` | Gửi tin nhắn vào channel | `teamId`, `channelId`, `message` |
| `list_chats` | Liệt kê chats | — |
| `get_chat_messages` | Đọc tin nhắn chat | `chatId`, `limit` |
| `send_chat_message` | Gửi tin nhắn chat | `chatId`, `message` |
| `list_meetings` | Liệt kê online meetings | `limit` |
| `get_meeting` | Chi tiết meeting | `meetingId` |
| `get_meeting_attendance` | Báo cáo điểm danh | `meetingId` |
| `get_meeting_transcripts` | Transcript meeting | `meetingId` |
| `search_messages` | Tìm kiếm tin nhắn Teams | `query`, `limit` |

---

## 10. Troubleshooting

| Lỗi | Nguyên nhân | Giải pháp |
|-----|-------------|-----------|
| "No authorization information present" | Token không có đủ permissions | Kiểm tra API permissions + admin consent |
| "Resource not found for segment 'v1.0'" | Graph SDK baseUrl conflict | Đảm bảo không set `baseUrl` trong SDK init |
| Infinite refresh_token loop | Token validation fail liên tục | Kiểm tra token cache, Graph API connectivity |
| "Error authenticating with resource" | User không có Teams license | Gán Microsoft 365 license cho user |
| "invalid_request: redirect_uri not valid" | Redirect URI không khớp | Kiểm tra URI trên Azure Portal vs code |
| Token expired (opaque) | Personal account token | Dùng work account; code đã handle opaque tokens |

---

## Checklist trước Go-live

- [ ] App Registration: single tenant, đúng permissions
- [ ] Admin consent đã được grant bởi Global Admin
- [ ] Redirect URI trên Azure khớp chính xác với production URL
- [ ] Authority dùng tenant-specific (không dùng `/common`)
- [ ] HTTPS enabled với valid certificate
- [ ] Client secret lưu trong Key Vault / secret manager
- [ ] Environment variables configured đúng
- [ ] Test thành công với work account có Teams license
- [ ] Logging và monitoring đã cấu hình
- [ ] Rate limiting enabled
- [ ] Backup plan cho secret rotation
