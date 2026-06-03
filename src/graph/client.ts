import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";

export interface AuthStatus {
  isAuthenticated: boolean;
  userPrincipalName?: string;
  displayName?: string;
}

interface CachedValidation {
  status: AuthStatus;
  expiresAt: number;
}

export class GraphService {
  private static instance: GraphService;
  private client: Client | undefined;
  private credential: ClientSecretCredential | undefined;
  private initialized = false;
  private tokenCache = new Map<string, CachedValidation>();
  private static TOKEN_CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes

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
    });

    this.initialized = true;
  }

  async getClient(userToken?: string): Promise<Client> {
    // Prefer user token (delegated permissions) over app token
    if (userToken) {
      return this.getClientWithToken(userToken);
    }
    this.initialize();
    if (!this.client) throw new Error("Graph client not initialized");
    return this.client;
  }

  getClientWithToken(token: string): Client {
    return Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => token,
      },
    });
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

  private decodeTokenPayload(token: string): Record<string, unknown> | null {
    try {
      const parts = token.split(".");
      if (parts.length !== 3) return null;
      const payload = Buffer.from(parts[1], "base64url").toString("utf-8");
      return JSON.parse(payload) as Record<string, unknown>;
    } catch {
      return null;
    }
  }

  private isTokenExpired(token: string): { expired: boolean; isOpaque: boolean } {
    const payload = this.decodeTokenPayload(token);
    if (!payload || typeof payload.exp !== "number") {
      // Opaque token (Microsoft personal accounts) - can't check expiry locally
      console.log("[Auth] Opaque token detected (not a decodable JWT) - will validate via Graph API");
      return { expired: false, isOpaque: true };
    }
    const now = Math.floor(Date.now() / 1000);
    const expired = payload.exp <= now;
    console.log(`[Auth] JWT token exp=${payload.exp} now=${now} diff=${payload.exp - now}s expired=${expired}`);
    return { expired, isOpaque: false };
  }

  async validateToken(token: string): Promise<AuthStatus> {
    // Check if token is expired by decoding JWT (skip for opaque tokens)
    const { expired, isOpaque } = this.isTokenExpired(token);
    if (expired) {
      console.log("[Auth] Token is expired (JWT exp check)");
      return { isAuthenticated: false };
    }

    // Check cache first
    const cacheKey = token.slice(-32); // Use last 32 chars as key to avoid storing full token
    const cached = this.tokenCache.get(cacheKey);
    if (cached && Date.now() < cached.expiresAt) {
      console.log("[Auth] Token validated from cache");
      return cached.status;
    }

    // Validate against Graph API using direct fetch (avoids SDK baseUrl issues)
    try {
      console.log("[Auth] Validating token against Microsoft Graph /me...");
      const response = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (!response.ok) {
        const errorBody = await response.text();
        console.error(`[Auth] Graph /me failed: status=${response.status} body=${errorBody}`);

        if (isOpaque) {
          console.log("[Auth] Opaque token rejected by Graph API");
          return { isAuthenticated: false };
        }

        // For JWT tokens, allow degraded mode if not expired
        const payload = this.decodeTokenPayload(token);
        if (payload && typeof payload.exp === "number") {
          const now = Math.floor(Date.now() / 1000);
          if (payload.exp > now) {
            console.log("[Auth] Graph call failed but JWT not expired - allowing access (degraded mode)");
            const fallbackStatus: AuthStatus = {
              isAuthenticated: true,
              userPrincipalName: (payload.upn as string) || (payload.preferred_username as string) || "unknown",
              displayName: (payload.name as string) || "Unknown User",
            };
            this.tokenCache.set(cacheKey, {
              status: fallbackStatus,
              expiresAt: Date.now() + 60_000,
            });
            return fallbackStatus;
          }
        }
        return { isAuthenticated: false };
      }

      const me = await response.json() as Record<string, unknown>;
      const status: AuthStatus = {
        isAuthenticated: true,
        userPrincipalName: me.userPrincipalName as string | undefined,
        displayName: me.displayName as string | undefined,
      };
      console.log(`[Auth] Token valid - user=${status.userPrincipalName}`);

      // Cache successful validation
      this.tokenCache.set(cacheKey, {
        status,
        expiresAt: Date.now() + GraphService.TOKEN_CACHE_TTL_MS,
      });

      return status;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.error(`[Auth] Token validation error: ${message}`);
      return { isAuthenticated: false };
    }
  }
}
