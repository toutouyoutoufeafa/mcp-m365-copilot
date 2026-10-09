import { MCPServer, createMCPServer } from "./mcp/server.js";
import { AuthConfig } from "./auth/entraAuth.js";
import { logger } from "./utils/logger.js";

function getAuthConfig(): AuthConfig {
  const clientId = process.env["ENTRA_CLIENT_ID"];
  const tenantId = process.env["ENTRA_TENANT_ID"];
  const redirectUri = process.env["ENTRA_REDIRECT_URI"] || "http://localhost:8400";

  if (!clientId || !tenantId) {
    throw new Error("ENTRA_CLIENT_ID and ENTRA_TENANT_ID environment variables are required");
  }

  return {
    clientId,
    tenantId,
    redirectUri,
    enableWam: process.env["ENTRA_ENABLE_WAM"] !== "false",
  };
}

function getWorkspaceRoot(): string {
  return process.env["WORKSPACE_ROOT"] || process.cwd();
}

async function main(): Promise<void> {
  const workspaceRoot = getWorkspaceRoot();
  const authConfig = getAuthConfig();

  logger.info("Starting MCP M365 Copilot Server", {
    workspaceRoot,
    clientId: authConfig.clientId,
    tenantId: authConfig.tenantId,
    enableWam: authConfig.enableWam,
  });

  const server = await createMCPServer(workspaceRoot, authConfig);

  process.on("SIGINT", async () => {
    logger.info("Received SIGINT, shutting down...");
    await server.stop();
    process.exit(0);
  });

  process.on("SIGTERM", async () => {
    logger.info("Received SIGTERM, shutting down...");
    await server.stop();
    process.exit(0);
  });

  try {
    await server.start();
  } catch (error) {
    logger.error("Server failed to start", { error: error instanceof Error ? error.message : String(error) });
    process.exit(1);
  }
}

main().catch((error) => {
  logger.error("Unhandled error in main", { error: error instanceof Error ? error.message : String(error) });
  process.exit(1);
});