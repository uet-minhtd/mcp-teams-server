import { describe, it, expect, vi, beforeEach } from "vitest";
import type { Request, Response } from "express";
import {
  oauthProtectedResourceHandler,
  oauthAuthorizationServerHandler,
  dcrRegisterHandler,
  oauthCallbackHandler,
  authorizeHandler,
  tokenExchangeHandler,
} from "../auth/endpoints.js";

function createMockReq(overrides: Partial<Request> = {}): Request {
  return {
    headers: {},
    query: {},
    body: {},
    get: vi.fn().mockReturnValue("localhost"),
    protocol: "http",
    ...overrides,
  } as unknown as Request;
}

function createMockRes(): Response {
  const res: Record<string, unknown> = {};
  res.statusCode = 200;
  res.json = vi.fn().mockReturnValue(res);
  res.status = vi.fn().mockReturnValue(res);
  res.send = vi.fn().mockReturnValue(res);
  res.redirect = vi.fn().mockReturnValue(res);
  return res as unknown as Response;
}

describe("OAuth Endpoints", () => {
  describe("oauthProtectedResourceHandler", () => {
    it("returns authorization_servers with current host", () => {
      const req = createMockReq();
      const res = createMockRes();
      const handler = oauthProtectedResourceHandler();

      handler(req, res);

      expect(res.json).toHaveBeenCalledWith({
        authorization_servers: ["http://localhost"],
      });
    });
  });

  describe("oauthAuthorizationServerHandler", () => {
    it("returns full OAuth metadata", () => {
      const req = createMockReq();
      const res = createMockRes();
      const handler = oauthAuthorizationServerHandler();

      handler(req, res);

      const arg = (res.json as ReturnType<typeof vi.fn>).mock.calls[0][0];
      expect(arg.issuer).toBe("http://localhost");
      expect(arg.authorization_endpoint).toBe("http://localhost/authorize");
      expect(arg.token_endpoint).toBe("http://localhost/token");
      expect(arg.registration_endpoint).toBe("http://localhost/register");
      expect(arg.response_types_supported).toEqual(["code"]);
      expect(arg.code_challenge_methods_supported).toEqual(["S256"]);
      expect(arg.grant_types_supported).toEqual([
        "authorization_code",
        "refresh_token",
      ]);
      expect(arg.scopes_supported).toContain("User.Read");
      expect(arg.scopes_supported).toContain("OnlineMeetings.Read");
      expect(arg.jwks_uri).toBe(
        "https://login.microsoftonline.com/common/discovery/v2.0/keys"
      );
    });
  });

  describe("dcrRegisterHandler", () => {
    it("returns DCR registration response", () => {
      const req = createMockReq();
      const res = createMockRes();
      const handler = dcrRegisterHandler();

      handler(req, res);

      const arg = (res.json as ReturnType<typeof vi.fn>).mock.calls[0][0];
      expect(arg.client_id).toBe("14d82eec-204b-4c2f-b7e8-296a70dab67e");
      expect(arg.redirect_uris).toContain("http://localhost");
      expect(arg.grant_types).toEqual(["authorization_code", "refresh_token"]);
      expect(arg.response_types).toEqual(["code"]);
      expect(arg.scope).toContain("User.Read");
      expect(arg.scope).toContain("OnlineMeetings.Read");
    });
  });

  describe("oauthCallbackHandler", () => {
    it("returns 400 when state or code is missing", () => {
      const req = createMockReq({ query: {} });
      const res = createMockRes();
      const handler = oauthCallbackHandler();

      handler(req, res);

      expect(res.status).toHaveBeenCalledWith(400);
      expect(res.send).toHaveBeenCalledWith("Missing state or code parameter");
    });

    it("returns 400 when state is unknown", () => {
      const req = createMockReq({
        query: { state: "unknown-state", code: "auth-code-123" },
      });
      const res = createMockRes();
      const handler = oauthCallbackHandler();

      handler(req, res);

      expect(res.status).toHaveBeenCalledWith(400);
      expect(res.send).toHaveBeenCalledWith("Invalid or expired state");
    });
  });

  describe("authorizeHandler", () => {
    it("redirects to Microsoft authorize endpoint", () => {
      const req = createMockReq({ query: {} });
      const res = createMockRes();
      const handler = authorizeHandler(8888);

      handler(req, res);

      expect(res.redirect).toHaveBeenCalledWith(
        302,
        expect.stringContaining(
          "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
        )
      );
    });

    it("includes default scopes when none provided", () => {
      const req = createMockReq({ query: {} });
      const res = createMockRes();
      const handler = authorizeHandler(8888);

      handler(req, res);

      const redirectUrl = (res.redirect as ReturnType<typeof vi.fn>).mock
        .calls[0][1] as string;
      expect(redirectUrl).toContain("scope=");
      expect(redirectUrl).toContain("User.Read");
      expect(redirectUrl).toContain("OnlineMeetings.Read");
    });
  });

  describe("tokenExchangeHandler", () => {
    it("returns 500 on fetch failure", async () => {
      const req = createMockReq({
        body: { code: "test-code", grant_type: "authorization_code" },
      });
      const res = createMockRes();
      const handler = tokenExchangeHandler(8888);

      global.fetch = vi.fn().mockRejectedValue(new Error("Network failure"));

      await handler(req, res);

      expect(res.status).toHaveBeenCalledWith(500);
      expect(res.json).toHaveBeenCalledWith({
        error: "token_exchange_failed",
        error_description: "Failed to exchange authorization code",
      });
    });
  });
});
