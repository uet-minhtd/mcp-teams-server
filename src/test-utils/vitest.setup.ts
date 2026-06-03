import { vi } from "vitest";

process.env.AZURE_CLIENT_ID = "test-client-id";
process.env.AZURE_CLIENT_SECRET = "test-client-secret";
process.env.AZURE_TENANT_ID = "test-tenant-id";

vi.mock("@azure/identity", () => ({
  ClientSecretCredential: vi.fn().mockImplementation(() => ({
    getToken: vi.fn().mockResolvedValue({ token: "mock-graph-token" }),
  })),
}));
