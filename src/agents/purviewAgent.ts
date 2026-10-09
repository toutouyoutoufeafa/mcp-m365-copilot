import { MCPTool, ToolResult } from "../mcp/types.js";
import { BaseAgent } from "./baseAgent.js";
import { EntraAuthManager } from "../auth/entraAuth.js";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { sanitizeText } from "../dlp/secretScanner.js";

export class PurviewAgent extends BaseAgent {
  private readonly purviewEndpoint: string;

  constructor(authManager: EntraAuthManager, purviewAccountName?: string) {
    const config = {
      type: "purview" as const,
      name: "Purview Agent",
      scopes: [
        "https://purview.azure.com/PurviewDataPolicies.Read.All",
        "https://purview.azure.com/PurviewCatalog.Read.All",
        "https://purview.azure.com/PurviewAccounts.Read.All",
      ],
      requiresApproval: false,
      maxTokensPerRequest: 4000,
      rateLimitPerMinute: 30,
    };
    super(config, authManager, "purview");
    
    this.purviewEndpoint = purviewAccountName 
      ? `https://${purviewAccountName}.purview.azure.com`
      : "";
  }

  override getTools(): MCPTool[] {
    const tools: MCPTool[] = [
      {
        name: "purview_list_accounts",
        description: "List accessible Purview accounts",
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "purview_list_data_sources",
        description: "List registered data sources in Purview",
        inputSchema: {
          type: "object",
          properties: {
            accountName: { type: "string", description: "Purview account name (optional if configured)" },
            type: { type: "string", description: "Filter by data source type (e.g., AzureBlobStorage, AzureSQLDatabase)" },
          },
        },
      },
      {
        name: "purview_list_assets",
        description: "Search/browse assets in Purview Data Catalog",
        inputSchema: {
          type: "object",
          properties: {
            accountName: { type: "string", description: "Purview account name" },
            query: { type: "string", description: "Search query (keywords)" },
            type: { type: "string", description: "Filter by asset type" },
            limit: { type: "number", description: "Max results", default: 50 },
            offset: { type: "number", description: "Pagination offset", default: 0 },
          },
          required: ["accountName"],
        },
      },
      {
        name: "purview_get_asset",
        description: "Get detailed asset information including schema, lineage, classification",
        inputSchema: {
          type: "object",
          properties: {
            accountName: { type: "string", description: "Purview account name" },
            assetGuid: { type: "string", description: "Asset GUID" },
          },
          required: ["accountName", "assetGuid"],
        },
      },
      {
        name: "purview_get_lineage",
        description: "Get data lineage for an asset",
        inputSchema: {
          type: "object",
          properties: {
            accountName: { type: "string", description: "Purview account name" },
            assetGuid: { type: "string", description: "Asset GUID" },
            direction: { type: "string", enum: ["upstream", "downstream", "both"], default: "both" },
            depth: { type: "number", default: 3 },
          },
          required: ["accountName", "assetGuid"],
        },
      },
      {
        name: "purview_list_scans",
        description: "List scan runs for a data source",
        inputSchema: {
          type: "object",
          properties: {
            accountName: { type: "string", description: "Purview account name" },
            dataSourceName: { type: "string", description: "Data source name" },
            status: { type: "string", enum: ["Running", "Succeeded", "Failed", "Cancelled"] },
          },
          required: ["accountName", "dataSourceName"],
        },
      },
      {
        name: "purview_get_scan_status",
        description: "Get detailed scan run status",
        inputSchema: {
          type: "object",
          properties: {
            accountName: { type: "string", description: "Purview account name" },
            dataSourceName: { type: "string", description: "Data source name" },
            scanId: { type: "string", description: "Scan run ID" },
          },
          required: ["accountName", "dataSourceName", "scanId"],
        },
      },
      {
        name: "purview_list_classifications",
        description: "List classification rules / sensitivity labels",
        inputSchema: {
          type: "object",
          properties: {
            accountName: { type: "string", description: "Purview account name" },
          },
          required: ["accountName"],
        },
      },
      {
        name: "purview_list_policies",
        description: "List data policies (access, retention, etc.)",
        inputSchema: {
          type: "object",
          properties: {
            accountName: { type: "string", description: "Purview account name" },
            type: { type: "string", enum: ["DataAccess", "DataRetention", "DataQuality"] },
          },
          required: ["accountName"],
        },
      },
      {
        name: "purview_search_assets",
        description: "Advanced search in Purview Catalog with filters",
        inputSchema: {
          type: "object",
          properties: {
            accountName: { type: "string", description: "Purview account name" },
            keywords: { type: "string", description: "Keywords to search" },
            filters: {
              type: "object",
              description: "Advanced filters",
              properties: {
                assetType: { type: "string" },
                classification: { type: "string" },
                glossaryTerm: { type: "string" },
                dataSource: { type: "string" },
                owner: { type: "string" },
              },
            },
            limit: { type: "number", default: 50 },
          },
          required: ["accountName"],
        },
      },
    ];

    return tools;
  }

  async executeTool(name: string, args: Record<string, unknown>): Promise<ToolResult> {
    const getArg = (key: string) => args[key];
    const accountName = (getArg("accountName") as string) || this.purviewEndpoint.replace("https://", "").replace(".purview.azure.com", "");

    if (!accountName && name !== "purview_list_accounts") {
      return {
        success: false,
        error: "Purview account name required (configure PURVIEW_ACCOUNT_NAME or pass accountName parameter)",
        auditEntry: {
          correlationId: createCorrelationId(),
          userIdHash: hashUserId(createCorrelationId()),
          tool: `purview.${name}`,
          action: name,
          result: "error",
          durationMs: 0,
          agentType: "purview",
        },
      };
    }

    const baseUrl = this.purviewEndpoint || `https://${accountName}.purview.azure.com`;

    switch (name) {
      case "purview_list_accounts":
        return this.executeWithGuardrails(name, args, () => this.listAccounts());
      case "purview_list_data_sources":
        return this.executeWithGuardrails(name, args, () => this.listDataSources(baseUrl));
      case "purview_list_assets":
        return this.executeWithGuardrails(name, args, () => this.listAssets(baseUrl, getArg("query") as string, getArg("type") as string, getArg("limit") as number, getArg("offset") as number));
      case "purview_get_asset":
        return this.executeWithGuardrails(name, args, () => this.getAsset(baseUrl, getArg("assetGuid") as string));
      case "purview_get_lineage":
        return this.executeWithGuardrails(name, args, () => this.getLineage(baseUrl, getArg("assetGuid") as string, getArg("direction") as string, getArg("depth") as number));
      case "purview_list_scans":
        return this.executeWithGuardrails(name, args, () => this.listScans(baseUrl, getArg("dataSourceName") as string, getArg("status") as string));
      case "purview_get_scan_status":
        return this.executeWithGuardrails(name, args, () => this.getScanStatus(baseUrl, getArg("dataSourceName") as string, getArg("scanId") as string));
      case "purview_list_classifications":
        return this.executeWithGuardrails(name, args, () => this.listClassifications(baseUrl));
      case "purview_list_policies":
        return this.executeWithGuardrails(name, args, () => this.listPolicies(baseUrl, getArg("type") as string));
      case "purview_search_assets":
        return this.executeWithGuardrails(name, args, () => this.searchAssets(baseUrl, getArg("keywords") as string, getArg("filters") as Record<string, unknown>, getArg("limit") as number));
      default:
        return {
          success: false,
          error: `Unknown tool: ${name}`,
          auditEntry: {
            correlationId: createCorrelationId(),
            userIdHash: hashUserId(createCorrelationId()),
            tool: `purview.${name}`,
            action: name,
            result: "error",
            durationMs: 0,
            agentType: "purview",
          },
        };
    }
  }

  private async listAccounts(): Promise<any> {
    const token = await this.authManager.getToken("purview");
    const response = await fetch("https://management.azure.com/providers/Microsoft.Purview/accounts?api-version=2023-07-01", {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Failed to list accounts: ${response.statusText}`);
    return response.json();
  }

  private async listDataSources(baseUrl: string): Promise<any> {
    const token = await this.authManager.getToken("purview");
    const response = await fetch(`${baseUrl}/datamap/api/datasources?api-version=2023-09-01`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Failed to list data sources: ${response.statusText}`);
    return response.json();
  }

  private async listAssets(baseUrl: string, query?: string, type?: string, limit: number = 50, offset: number = 0): Promise<any> {
    const token = await this.authManager.getToken("purview");
    const params = new URLSearchParams({
      limit: limit.toString(),
      offset: offset.toString(),
    });
    if (query) params.append("keywords", query);
    if (type) params.append("type", type);

    const response = await fetch(`${baseUrl}/datamap/api/search/query?${params}`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ keywords: query || "" }),
    });
    if (!response.ok) throw new Error(`Failed to list assets: ${response.statusText}`);
    return response.json();
  }

  private async getAsset(baseUrl: string, assetGuid: string): Promise<any> {
    const token = await this.authManager.getToken("purview");
    const response = await fetch(`${baseUrl}/datamap/api/entities/guid/${assetGuid}?api-version=2023-09-01`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Failed to get asset: ${response.statusText}`);
    return response.json();
  }

  private async getLineage(baseUrl: string, assetGuid: string, direction: string = "both", depth: number = 3): Promise<any> {
    const token = await this.authManager.getToken("purview");
    const response = await fetch(`${baseUrl}/datamap/api/lineage/${assetGuid}?direction=${direction}&depth=${depth}&api-version=2023-09-01`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Failed to get lineage: ${response.statusText}`);
    return response.json();
  }

  private async listScans(baseUrl: string, dataSourceName: string, status?: string): Promise<any> {
    const token = await this.authManager.getToken("purview");
    const params = new URLSearchParams();
    if (status) params.append("status", status);
    const response = await fetch(`${baseUrl}/datamap/api/datasources/${encodeURIComponent(dataSourceName)}/scans?${params}&api-version=2023-09-01`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Failed to list scans: ${response.statusText}`);
    return response.json();
  }

  private async getScanStatus(baseUrl: string, dataSourceName: string, scanId: string): Promise<any> {
    const token = await this.authManager.getToken("purview");
    const response = await fetch(`${baseUrl}/datamap/api/datasources/${encodeURIComponent(dataSourceName)}/scans/${scanId}?api-version=2023-09-01`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Failed to get scan status: ${response.statusText}`);
    return response.json();
  }

  private async listClassifications(baseUrl: string): Promise<any> {
    const token = await this.authManager.getToken("purview");
    const response = await fetch(`${baseUrl}/datamap/api/classifications?api-version=2023-09-01`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Failed to list classifications: ${response.statusText}`);
    return response.json();
  }

  private async listPolicies(baseUrl: string, type?: string): Promise<any> {
    const token = await this.authManager.getToken("purview");
    const params = new URLSearchParams();
    if (type) params.append("type", type);
    const response = await fetch(`${baseUrl}/policies?${params}&api-version=2023-07-01`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Failed to list policies: ${response.statusText}`);
    return response.json();
  }

  private async searchAssets(baseUrl: string, keywords: string, filters?: Record<string, unknown>, limit: number = 50): Promise<any> {
    const token = await this.authManager.getToken("purview");
    const body: Record<string, unknown> = { keywords, limit };
    if (filters) body["filters"] = filters;

    const response = await fetch(`${baseUrl}/datamap/api/search/query?api-version=2023-09-01`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!response.ok) throw new Error(`Failed to search assets: ${response.statusText}`);
    return response.json();
  }
}