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
