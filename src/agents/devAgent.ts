import { MCPTool, ToolResult } from "../mcp/types.js";
import { BaseAgent } from "./baseAgent.js";
import { EntraAuthManager } from "../auth/entraAuth.js";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { scanForSecrets, sanitizeText } from "../dlp/secretScanner.js";
import { readFile as fsReadFile, writeFile as fsWriteFile, mkdir, readdir, stat } from "fs/promises";
import { join, resolve, relative, extname } from "path";

const ALLOWED_EXTENSIONS = [".ps1", ".py", ".tf", ".json", ".yaml", ".yml", ".md", ".txt", ".ts", ".js"];
const MAX_FILE_SIZE = 1024 * 1024;

export class DevAgent extends BaseAgent {
  private readonly workspaceRoot: string;

  constructor(authManager: EntraAuthManager, workspaceRoot: string) {
    const config = {
      type: "dev" as const,
      name: "Development Agent",
      scopes: ["User.Read", "Files.Read"],
      allowedPaths: ["Standards/", "Runbooks/", "Templates/", "Architecture/", "src/", "tests/"],
      requiresApproval: true,
      maxTokensPerRequest: 8000,
      rateLimitPerMinute: 30,
    };
    super(config, authManager, "dev");
    this.workspaceRoot = resolve(workspaceRoot);
  }

  override getTools(): MCPTool[] {
    return [
      {
        name: "dev_read_file",
        description: "Read a file from the local workspace (allowed paths only)",
        inputSchema: {
          type: "object",
          properties: {
            path: { type: "string", description: "Relative path from workspace root" },
          },
          required: ["path"],
        },
      },
      {
        name: "dev_write_file",
        description: "Write a file to the local workspace (requires human approval preview)",
        inputSchema: {
          type: "object",
          properties: {
            path: { type: "string", description: "Relative path from workspace root" },
            content: { type: "string", description: "File content" },
            previewOnly: { type: "boolean", description: "If true, only show preview without writing", default: true },
          },
          required: ["path", "content"],
        },
      },
      {
        name: "dev_list_directory",
        description: "List directory contents in the workspace",
        inputSchema: {
          type: "object",
          properties: {
            path: { type: "string", description: "Relative path from workspace root", default: "." },
          },
        },
      },
      {
        name: "dev_generate_powershell",
        description: "Generate PowerShell script for Microsoft 365/Intune/Graph operations",
        inputSchema: {
          type: "object",
          properties: {
            task: { type: "string", description: "Description of the task to automate" },
            target: { type: "string", enum: ["graph", "intune", "teams", "sharepoint", "exchange"], description: "Target service" },
            parameters: { type: "object", description: "Parameters for the script" },
          },
          required: ["task", "target"],
        },
      },
      {
        name: "dev_generate_python",
        description: "Generate Python script for Microsoft 365/Graph operations",
        inputSchema: {
          type: "object",
          properties: {
            task: { type: "string", description: "Description of the task to automate" },
            target: { type: "string", enum: ["graph", "intune", "teams", "sharepoint", "exchange"], description: "Target service" },
            parameters: { type: "object", description: "Parameters for the script" },
          },
          required: ["task", "target"],
        },
      },
      {
        name: "dev_generate_terraform",
        description: "Generate Terraform configuration for Microsoft 365/Intune resources",
        inputSchema: {
          type: "object",
          properties: {
            resources: { type: "array", items: { type: "string" }, description: "List of resources to manage" },
            environment: { type: "string", enum: ["dev", "test", "prod"], description: "Target environment" },
          },
          required: ["resources"],
        },
      },
    ];
  }

  async executeTool(name: string, args: Record<string, unknown>): Promise<ToolResult> {
    const getArg = (key: string) => args[key];

    switch (name) {
      case "dev_read_file":
        return this.executeWithGuardrails(name, args, () => this.readFile(getArg("path") as string));
      case "dev_write_file":
        return this.executeWithGuardrails(name, args, () =>
          this.writeFile(getArg("path") as string, getArg("content") as string, (getArg("previewOnly") as boolean) ?? true)
        );
      case "dev_list_directory":
        return this.executeWithGuardrails(name, args, () => this.listDirectory((getArg("path") as string) || "."));
      case "dev_generate_powershell":
        return this.executeWithGuardrails(name, args, () =>
          this.generatePowerShell(getArg("task") as string, getArg("target") as string, getArg("parameters") as Record<string, unknown>)
        );
      case "dev_generate_python":
        return this.executeWithGuardrails(name, args, () =>
          this.generatePython(getArg("task") as string, getArg("target") as string, getArg("parameters") as Record<string, unknown>)
        );
      case "dev_generate_terraform":
        return this.executeWithGuardrails(name, args, () =>
          this.generateTerraform((getArg("resources") as string[]) || [], getArg("environment") as string)
        );
      default:
        return {
          success: false,
          error: `Unknown tool: ${name}`,
          auditEntry: {
            correlationId: createCorrelationId(),
            userIdHash: hashUserId(createCorrelationId()),
            tool: `dev.${name}`,
            action: name,
            result: "error",
            durationMs: 0,
            agentType: "dev",
          },
        };
    }
  }

  private async readFile(path: string): Promise<string> {
    const fullPath = this.resolveAndValidatePath(path);
    const fileStat = await stat(fullPath);
    if (fileStat.size > MAX_FILE_SIZE) {
      throw new Error(`File too large: ${fileStat.size} bytes (max ${MAX_FILE_SIZE})`);
    }
    const content = await fsReadFile(fullPath, "utf-8");
    return sanitizeText(content);
  }

  private async writeFile(path: string, content: string, previewOnly: boolean = true): Promise<{ preview: string; written: boolean }> {
    const fullPath = this.resolveAndValidatePath(path);
    const secrets = scanForSecrets(content);
    if (secrets.some((s) => s.severity === "critical")) {
      throw new Error("Critical secrets detected in content, write blocked");
    }

    if (previewOnly) {
      return { preview: sanitizeText(content), written: false };
    }

    await mkdir(join(fullPath, ".."), { recursive: true });
    await fsWriteFile(fullPath, content, "utf-8");
    return { preview: sanitizeText(content), written: true };
  }

  private async listDirectory(path: string): Promise<{ path: string; name: string; isDirectory: boolean; size?: number }[]> {
    const fullPath = this.resolveAndValidatePath(path);
    const entries = await readdir(fullPath, { withFileTypes: true });
    const results = await Promise.all(
      entries.map(async (entry) => ({
        path: join(path, entry.name),
        name: entry.name,
        isDirectory: entry.isDirectory(),
        size: entry.isFile() ? (await stat(join(fullPath, entry.name))).size : undefined,
      }))
    );
    return results;
  }

  private async generatePowerShell(task: string, target: string, parameters: Record<string, unknown> = {}): Promise<string> {
    const sanitizedTask = sanitizeText(task);
    const sanitizedParams = JSON.stringify(parameters, null, 2);

    const templates = {
      graph: `# PowerShell Script for Microsoft Graph
# Task: ${sanitizedTask}
# Generated: ${new Date().toISOString()}
# Parameters: ${sanitizedParams}

# Requires: Microsoft.Graph module
# Install-Module Microsoft.Graph -Scope CurrentUser

Import-Module Microsoft.Graph

# Connect with delegated permissions (User.Read, Files.Read)
Connect-MgGraph -Scopes "User.Read", "Files.Read"

# TODO: Implement Graph operations for: ${sanitizedTask}
# Example:
# $user = Get-MgUser -UserId "user@domain.com"
# Write-Output $user

Disconnect-MgGraph
`,
      intune: `# PowerShell Script for Intune
# Task: ${sanitizedTask}
# Generated: ${new Date().toISOString()}
# Parameters: ${sanitizedParams}

# Requires: Microsoft.Graph.Intune module
# Install-Module Microsoft.Graph.Intune -Scope CurrentUser

Import-Module Microsoft.Graph.Intune

Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "DeviceManagementConfiguration.Read.All"

# TODO: Implement Intune operations for: ${sanitizedTask}

Disconnect-MgGraph
`,
      teams: `# PowerShell Script for Teams
# Task: ${sanitizedTask}
# Generated: ${new Date().toISOString()}
# Parameters: ${sanitizedParams}

# Requires: MicrosoftTeams module
# Install-Module MicrosoftTeams -Scope CurrentUser

Import-Module MicrosoftTeams

# Connect with delegated permissions
Connect-MicrosoftTeams

# TODO: Implement Teams operations for: ${sanitizedTask}

Disconnect-MicrosoftTeams
`,
      sharepoint: `# PowerShell Script for SharePoint
# Task: ${sanitizedTask}
# Generated: ${new Date().toISOString()}
# Parameters: ${sanitizedParams}

# Requires: PnP.PowerShell module
# Install-Module PnP.PowerShell -Scope CurrentUser

Import-Module PnP.PowerShell

# Connect with interactive login
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/site" -Interactive

# TODO: Implement SharePoint operations for: ${sanitizedTask}

Disconnect-PnPOnline
`,
      exchange: `# PowerShell Script for Exchange Online
# Task: ${sanitizedTask}
# Generated: ${new Date().toISOString()}
# Parameters: ${sanitizedParams}

# Requires: ExchangeOnlineManagement module
# Install-Module ExchangeOnlineManagement -Scope CurrentUser

Import-Module ExchangeOnlineManagement

# Connect with delegated permissions
Connect-ExchangeOnline -UserPrincipalName "admin@domain.com"

# TODO: Implement Exchange operations for: ${sanitizedTask}

Disconnect-ExchangeOnline
`,
    } as const;

    const validTargets = ["graph", "intune", "teams", "sharepoint", "exchange"] as const;
    const validTarget = validTargets.includes(target as any) ? target : "graph";
    return templates[validTarget as keyof typeof templates] ?? templates.graph;
  }

  private async generatePython(task: string, target: string, parameters: Record<string, unknown> = {}): Promise<string> {
    const sanitizedTask = sanitizeText(task);
    const sanitizedParams = JSON.stringify(parameters, null, 2);

    const templates = {
      graph: `#!/usr/bin/env python3
"""
Python script for Microsoft Graph
Task: ${sanitizedTask}
Generated: ${new Date().toISOString()}
Parameters: ${sanitizedParams}
"""

from azure.identity import InteractiveBrowserCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilder

# Scopes: User.Read, Files.Read
SCOPES = ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/Files.Read"]

async def main():
    credential = InteractiveBrowserCredential()
    client = GraphServiceClient(credential, scopes=SCOPES)

    # TODO: Implement Graph operations for: ${sanitizedTask}
    # Example:
    # user = await client.me.get()
    # print(user.display_name)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
`,
      intune: `#!/usr/bin/env python3
"""
Python script for Intune
Task: ${sanitizedTask}
Generated: ${new Date().toISOString()}
Parameters: ${sanitizedParams}
"""

from azure.identity import InteractiveBrowserCredential
from msgraph import GraphServiceClient

# Scopes for Intune
SCOPES = [
    "https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All",
    "https://graph.microsoft.com/DeviceManagementConfiguration.Read.All",
]

async def main():
    credential = InteractiveBrowserCredential()
    client = GraphServiceClient(credential, scopes=SCOPES)

    # TODO: Implement Intune operations for: ${sanitizedTask}

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
`,
      teams: `#!/usr/bin/env python3
"""
Python script for Teams
Task: ${sanitizedTask}
Generated: ${new Date().toISOString()}
Parameters: ${sanitizedParams}
"""

from azure.identity import InteractiveBrowserCredential
from msgraph import GraphServiceClient

# Scopes for Teams
SCOPES = [
    "https://graph.microsoft.com/Team.ReadBasic.All",
    "https://graph.microsoft.com/Channel.ReadBasic.All",
]

async def main():
    credential = InteractiveBrowserCredential()
    client = GraphServiceClient(credential, scopes=SCOPES)

    # TODO: Implement Teams operations for: ${sanitizedTask}

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
`,
    } as const;

    const validPyTargets = ["graph", "intune", "teams"] as const;
    const validPyTarget = validPyTargets.includes(target as any) ? target : "graph";
    return templates[validPyTarget as keyof typeof templates] ?? templates.graph;
  }

  private async generateTerraform(resources: string[], environment: string = "dev"): Promise<string> {
    const sanitizedResources = resources.map((r) => sanitizeText(r)).join(", ");
    return `# Terraform configuration for Microsoft 365/Intune
# Resources: ${sanitizedResources}
# Environment: ${environment}
# Generated: ${new Date().toISOString()}

terraform {
  required_version = ">= 1.5"
  required_providers {
    azuread = {
      source  = "hashicorp/azuread"
      version = "~> 2.40"
    }
    msgraph = {
      source  = "microsoft/msgraph"
      version = "~> 1.0"
    }
  }
}

provider "azuread" {
  # Configured via AZURE_SUBSCRIPTION_ID, etc.
}

provider "msgraph" {
  # Uses Azure CLI or InteractiveBrowserCredential
}

# TODO: Define resources for: ${sanitizedResources}
# Example:
# resource "msgraph_user" "example" {
#   display_name = "Example User"
#   mail_nickname = "example"
#   user_principal_name = "example@domain.com"
# }
`;
  }

  private resolveAndValidatePath(path: string): string {
    const normalized = path.replace(/\\/g, "/");
    const fullPath = resolve(this.workspaceRoot, normalized);

    if (!fullPath.startsWith(this.workspaceRoot)) {
      throw new Error(`Path traversal attempt blocked: ${path}`);
    }

    const relPath = relative(this.workspaceRoot, fullPath);
    const allowed = this.config.allowedPaths?.some((allowedPath) => relPath.startsWith(allowedPath)) ?? false;
    if (!allowed && relPath !== "." && !relPath.startsWith("..")) {
      const ext = extname(fullPath).toLowerCase();
      if (!ALLOWED_EXTENSIONS.includes(ext) && !fullPath.endsWith("package.json") && !fullPath.endsWith("tsconfig.json")) {
        throw new Error(`Path not in allowed directories: ${relPath}`);
      }
    }

    return fullPath;
  }
}