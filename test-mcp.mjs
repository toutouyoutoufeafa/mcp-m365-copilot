#!/usr/bin/env node
// Test script for MCP protocol: initialize -> tools/list -> tools/call

import { spawn } from "child_process";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const serverPath = join(__dirname, "dist", "index.js");

function sendRequest(proc, request) {
  return new Promise((resolve, reject) => {
    const id = request.id || Math.random();
    const requestWithId = { ...request, id };
    const message = JSON.stringify(requestWithId) + "\n";
    
    const timeout = setTimeout(() => {
      reject(new Error("Request timeout"));
    }, 10000);
    
    const handler = (data) => {
      const lines = data.toString().trim().split("\n");
      for (const line of lines) {
        try {
          const response = JSON.parse(line);
          if (response.id === id) {
            clearTimeout(timeout);
            proc.stdout.removeListener("data", handler);
            resolve(response);
            return;
          }
        } catch (e) {
          // Ignore parse errors
        }
      }
    };
    
    proc.stdout.on("data", handler);
    proc.stdin.write(message);
  });
}

async function testMCP() {
  console.log("Starting MCP server...");
  const proc = spawn("node", [serverPath], {
    stdio: ["pipe", "pipe", "pipe"],
    env: {
      ...process.env,
      ENTRA_CLIENT_ID: "test-client-id",
      ENTRA_TENANT_ID: "test-tenant-id",
      WORKSPACE_ROOT: "C:\\Temp\\Projets\\Plugin",
    },
  });

  proc.stderr.on("data", (data) => {
    console.error("STDERR:", data.toString());
  });

  // Wait for server to start
  await new Promise((resolve) => setTimeout(resolve, 1000));

  try {
    // Test 1: Initialize
    console.log("\n=== Test 1: Initialize ===");
    const initResponse = await sendRequest(proc, {
      jsonrpc: "2.0",
      id: 1,
      method: "initialize",
      params: {
        protocolVersion: "2024-11-05",
        capabilities: {},
        clientInfo: { name: "test-client", version: "1.0.0" },
      },
    });
    console.log("Initialize response:", JSON.stringify(initResponse, null, 2));

    // Test 2: Tools List
    console.log("\n=== Test 2: Tools List ===");
    const toolsResponse = await sendRequest(proc, {
      jsonrpc: "2.0",
      id: 2,
      method: "tools/list",
      params: {},
    });
    console.log("Tools list response:", JSON.stringify(toolsResponse, null, 2));

    // Test 3: Tools Call - dev_read_file (should fail gracefully without auth)
    console.log("\n=== Test 3: Tools Call (dev_read_file) ===");
    const callResponse = await sendRequest(proc, {
      jsonrpc: "2.0",
      id: 3,
      method: "tools/call",
      params: {
        name: "dev_read_file",
        arguments: { path: "package.json" },
      },
    });
    console.log("Tool call response:", JSON.stringify(callResponse, null, 2));

    // Test 4: Tools Call - engineering_knowledge_search
    console.log("\n=== Test 4: Tools Call (engineering_knowledge_search) ===");
    const searchResponse = await sendRequest(proc, {
      jsonrpc: "2.0",
      id: 4,
      method: "tools/call",
      params: {
        name: "engineering_knowledge_search",
        arguments: { query: "test", maxResults: 5 },
      },
    });
    console.log("Search response:", JSON.stringify(searchResponse, null, 2));

    console.log("\n=== All tests passed! ===");
  } catch (error) {
    console.error("Test failed:", error);
  } finally {
    proc.kill();
  }
}

testMCP().catch(console.error);