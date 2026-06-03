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

const MICROSOFT_CLIENT_ID = process.env.AZURE_CLIENT_ID || "14d82eec-204b-4c2f-b7e8-296a70dab67e";

const CALLBACK_PROXY_TIMEOUT_MS = 5 * 60 * 1000;
const MAX_PENDING_CALLBACKS = 100;
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
    console.log("[OAuth] GET /.well-known/oauth-protected-resource");
    const host = _req.get("host") || "localhost";
    const proto = _req.protocol;
    const baseUrl = `${proto}://${host}`;
    res.json({ authorization_servers: [baseUrl] });
  };
}

export function oauthAuthorizationServerHandler() {
  return (req: Request, res: Response) => {
    console.log("[OAuth] GET /.well-known/oauth-authorization-server");
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
    console.log("[OAuth] GET /authorize - Starting authorization flow");
    console.log(`[OAuth]   client_id=${req.query.client_id || "none"}`);
    console.log(`[OAuth]   redirect_uri=${req.query.redirect_uri || "none"}`);
    console.log(`[OAuth]   scope=${req.query.scope || "default"}`);
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
      if (pendingCallbacks.size >= MAX_PENDING_CALLBACKS) {
        console.warn("[OAuth] Too many pending callbacks, rejecting");
        res.status(429).send("Too many pending authorization requests");
        return;
      }
      pendingCallbacks.set(state, { redirectUri: ru, timestamp: Date.now() });
      const proxyRu = `http://localhost:${port}/callback`;
      redirectUrl.searchParams.set("redirect_uri", proxyRu);
      console.log(`[OAuth]   Stored callback state=${state.substring(0, 8)}... redirecting to Microsoft`);
      console.log(`[OAuth]   proxy redirect_uri=${proxyRu}`);
    }
    res.redirect(302, redirectUrl.toString());
  };
}

export function oauthCallbackHandler() {
  return (req: Request, res: Response) => {
    const state = req.query.state as string | undefined;
    const code = req.query.code as string | undefined;
    const error = req.query.error as string | undefined;
    const errorDescription = req.query.error_description as string | undefined;

    console.log("[OAuth] GET /callback - Microsoft callback received");
    console.log(`[OAuth]   code=${code ? code.substring(0, 20) + "..." : "missing"}`);
    console.log(`[OAuth]   state=${state ? state.substring(0, 8) + "..." : "missing"}`);
    console.log(`[OAuth]   error=${error || "none"}`);
    console.log(`[OAuth]   error_description=${errorDescription || "none"}`);
    console.log(`[OAuth]   Full query params: ${JSON.stringify(req.query)}`);

    if (error) {
      console.error(`[OAuth] Microsoft returned error: ${error} - ${errorDescription}`);
      res.status(400).send(`OAuth error: ${error} - ${errorDescription}`);
      return;
    }

    if (!state || !code) {
      console.warn("[OAuth] Callback missing state or code");
      res.status(400).send("Missing state or code parameter");
      return;
    }
    const pending = pendingCallbacks.get(state);
    if (!pending) {
      console.warn(`[OAuth] Callback state not found or expired: ${state.substring(0, 8)}...`);
      console.warn(`[OAuth]   Pending callbacks count: ${pendingCallbacks.size}`);
      res.status(400).send("Invalid or expired state");
      return;
    }
    pendingCallbacks.delete(state);
    const redirectTo = new URL(pending.redirectUri);
    redirectTo.searchParams.set("code", code);
    redirectTo.searchParams.set("state", state);
    console.log(`[OAuth] Callback success - redirecting to client: ${pending.redirectUri.substring(0, 60)}...`);
    res.redirect(302, redirectTo.toString());
  };
}

export function dcrRegisterHandler() {
  return (_req: Request, res: Response) => {
    console.log("[OAuth] POST /register - Dynamic Client Registration");
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
    const grantType = req.body?.grant_type || "unknown";
    console.log(`[OAuth] POST /token - grant_type=${grantType}`);
    try {
      const body = new URLSearchParams();
      for (const [key, value] of Object.entries(req.body)) {
        if (typeof value === "string") {
          body.set(key, value);
        }
      }
      body.set("redirect_uri", `http://localhost:${port}/callback`);
      if (!body.has("scope")) {
        body.set("scope", DCR_SCOPES.join(" "));
      }
      // Ensure client credentials are included for confidential app
      if (!body.has("client_id")) {
        body.set("client_id", MICROSOFT_CLIENT_ID);
      }
      if (!body.has("client_secret") && process.env.AZURE_CLIENT_SECRET) {
        body.set("client_secret", process.env.AZURE_CLIENT_SECRET);
      }

      console.log(`[OAuth]   Exchanging with Microsoft token endpoint...`);
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
      console.log(`[OAuth]   Token response status=${response.status} has_access_token=${"access_token" in data}`);
      if (data.access_token) {
        const tokenStr = data.access_token as string;
        try {
          const parts = tokenStr.split(".");
          if (parts.length === 3) {
            const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf-8"));
            console.log(`[OAuth]   New token exp=${payload.exp} (expires in ${payload.exp - Math.floor(Date.now() / 1000)}s)`);
          }
        } catch { /* ignore decode error */ }
      }
      if (data.error) {
        console.warn(`[OAuth]   Token error: ${data.error} - ${data.error_description}`);
      }
      res.status(response.status).json(data);
    } catch (error) {
      console.error("[OAuth] Token exchange failed:", error);
      res.status(500).json({
        error: "token_exchange_failed",
        error_description: "Failed to exchange authorization code",
      });
    }
  };
}
