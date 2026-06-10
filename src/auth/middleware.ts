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
    console.log("Authentication: "+ authHeader);

    if (!authHeader?.startsWith("Bearer ")) {
      console.log(`[Auth Middleware] No bearer token - ${req.method} ${req.path}`);
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
    console.log(`[Auth Middleware] Validating token for ${req.method} ${req.path} (token length=${token.length})`);
    const graphService = GraphService.getInstance();
    const status = await graphService.validateToken(token);

    if (!status.isAuthenticated) {
      console.warn(`[Auth Middleware] Token rejected - ${req.method} ${req.path}`);
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

    console.log(`[Auth Middleware] Authorized: ${status.userPrincipalName} - ${req.method} ${req.path}`);
    // Auth context is stored on the request and passed to session server on creation
    (req as unknown as Record<string, unknown>).auth = {
      token,
      userId: status.userPrincipalName ?? "",
      displayName: status.displayName ?? "",
      userPrincipalName: status.userPrincipalName ?? "",
    } satisfies AuthContext;

    next();
  };
}
