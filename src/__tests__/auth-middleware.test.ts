import { describe, it, expect, vi, beforeEach } from "vitest";
import type { Request, Response, NextFunction } from "express";
import { requireBearerAuth } from "../auth/middleware.js";

const mockValidateToken = vi.fn();
vi.mock("../graph/client.js", () => ({
  GraphService: {
    getInstance: () => ({
      validateToken: mockValidateToken,
    }),
  },
}));

function createMockReq(overrides: Partial<Request> = {}): Request {
  return {
    headers: {},
    get: vi.fn().mockReturnValue("localhost"),
    protocol: "http",
    ...overrides,
  } as unknown as Request;
}

function createMockRes(): Response {
  const res: Record<string, unknown> = {};
  res.statusCode = 200;
  res.headers = {};
  res._headers = {};
  res.status = vi.fn().mockReturnValue(res);
  res.json = vi.fn().mockReturnValue(res);
  res.setHeader = vi.fn().mockImplementation((name: string, value: string) => {
    (res._headers as Record<string, string>)[name] = value;
    return res;
  });
  return res as unknown as Response;
}

describe("requireBearerAuth middleware", () => {
  let middleware: ReturnType<typeof requireBearerAuth>;

  beforeEach(() => {
    vi.clearAllMocks();
    middleware = requireBearerAuth();
  });

  it("returns 401 when no Authorization header", async () => {
    const req = createMockReq();
    const res = createMockRes();
    const next = vi.fn();

    await middleware(req, res, next);

    expect(res.status).toHaveBeenCalledWith(401);
    expect(res.json).toHaveBeenCalledWith(
      expect.objectContaining({
        error: expect.objectContaining({
          code: -32001,
          message: "Unauthorized: Bearer token required",
        }),
      })
    );
    expect(next).not.toHaveBeenCalled();
  });

  it("returns 401 when Authorization is not Bearer", async () => {
    const req = createMockReq({
      headers: { authorization: "Basic abc123" },
    });
    const res = createMockRes();
    const next = vi.fn();

    await middleware(req, res, next);

    expect(res.status).toHaveBeenCalledWith(401);
    expect(next).not.toHaveBeenCalled();
  });

  it("returns 401 when token is invalid", async () => {
    mockValidateToken.mockResolvedValue({ isAuthenticated: false });

    const req = createMockReq({
      headers: { authorization: "Bearer invalid-token" },
    });
    const res = createMockRes();
    const next = vi.fn();

    await middleware(req, res, next);

    expect(res.status).toHaveBeenCalledWith(401);
    expect(res.json).toHaveBeenCalledWith(
      expect.objectContaining({
        error: expect.objectContaining({
          message: "Unauthorized: Invalid or expired token",
        }),
      })
    );
    expect(next).not.toHaveBeenCalled();
  });

  it("calls next() and sets auth context for valid token", async () => {
    mockValidateToken.mockResolvedValue({
      isAuthenticated: true,
      userPrincipalName: "user@test.com",
      displayName: "Test User",
    });

    const req = createMockReq({
      headers: { authorization: "Bearer valid-token" },
    });
    const res = createMockRes();
    const next = vi.fn();

    await middleware(req, res, next);

    expect(next).toHaveBeenCalled();
    expect(mockValidateToken).toHaveBeenCalledWith("valid-token");

    const auth = (req as unknown as Record<string, unknown>).auth as {
      token: string;
      userId: string;
      displayName: string;
      userPrincipalName: string;
    };
    expect(auth.token).toBe("valid-token");
    expect(auth.userPrincipalName).toBe("user@test.com");
    expect(auth.displayName).toBe("Test User");
  });

  it("sets WWW-Authenticate header on 401", async () => {
    const req = createMockReq();
    const res = createMockRes();
    const next = vi.fn();

    await middleware(req, res, next);

    expect(res.setHeader).toHaveBeenCalledWith(
      "WWW-Authenticate",
      expect.stringContaining("resource_metadata=")
    );
  });
});
