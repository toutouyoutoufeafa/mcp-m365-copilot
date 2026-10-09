export interface ToolResult {
  success: boolean;
  data?: unknown;
  error?: string;
  auditEntry: Omit<AuditLogEntry, "timestamp">;
}

export interface MCPRequest {
  jsonrpc: "2.0";
  id: string | number;
  method: string;
  params?: Record<string, unknown>;
}

export interface MCPResponse {
  jsonrpc: "2.0";
  id: string | number;
  result?: Record<string, unknown>;
  error?: {
    code: number;
    message: string;
    data?: unknown;
  };
}

export interface MCPTool {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: Record<string, unknown>;
    required?: string[];
  };
}

export interface MCPInitializeParams {
  protocolVersion: string;
  capabilities: Record<string, unknown>;
  clientInfo: {
    name: string;
    version: string;
  };
}

export interface MCPInitializeResult {
  protocolVersion: string;
  capabilities: {
    tools: { listChanged: boolean };
    resources?: { subscribe: boolean; listChanged: boolean };
    prompts?: { listChanged: boolean };
  };
  serverInfo: {
    name: string;
    version: string;
  };
}

export type AgentType = "dev" | "knowledge" | "collaboration" | "purview" | "autopilot" | "defender" | "m365admin" | "exchange";

export interface AgentConfig {
  type: AgentType;
  name: string;
  scopes: string[];
  allowedPaths?: string[];
  allowlist?: string[];
  requiresApproval: boolean;
  maxTokensPerRequest: number;
  rateLimitPerMinute: number;
}

export const AGENT_CONFIGS: Record<AgentType, AgentConfig> = {
  dev: {
    type: "dev",
    name: "Development Agent",
    scopes: ["User.Read", "Files.Read"],
    allowedPaths: ["Standards/", "Runbooks/", "Templates/", "Architecture/", "src/", "tests/"],
    requiresApproval: true,
    maxTokensPerRequest: 8000,
    rateLimitPerMinute: 30,
  },
  knowledge: {
    type: "knowledge",
    name: "Knowledge Agent",
    scopes: ["User.Read", "Files.Read"],
    allowlist: ["Standards/", "Runbooks/", "Templates/", "Architecture/"],
    requiresApproval: false,
    maxTokensPerRequest: 4000,
    rateLimitPerMinute: 60,
  },
  collaboration: {
    type: "collaboration",
    name: "Collaboration Agent",
    scopes: ["User.Read", "Files.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All"],
    requiresApproval: true,
    maxTokensPerRequest: 4000,
    rateLimitPerMinute: 20,
  },
  purview: {
    type: "purview",
    name: "Purview Agent",
    scopes: [
      "https://purview.azure.com/PurviewDataPolicies.Read.All",
      "https://purview.azure.com/PurviewCatalog.Read.All",
      "https://purview.azure.com/PurviewAccounts.Read.All",
    ],
    requiresApproval: false,
    maxTokensPerRequest: 4000,
    rateLimitPerMinute: 30,
  },
  autopilot: {
    type: "autopilot",
    name: "Autopilot Agent",
    scopes: [
      "https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All",
      "https://graph.microsoft.com/DeviceManagementManagedDevices.ReadWrite.All",
      "https://graph.microsoft.com/DeviceManagementServiceConfig.Read.All",
      "https://graph.microsoft.com/DeviceManagementServiceConfig.ReadWrite.All",
      "https://graph.microsoft.com/DeviceManagementConfiguration.Read.All",
      "https://graph.microsoft.com/DeviceManagementConfiguration.ReadWrite.All",
    ],
    requiresApproval: true,
    maxTokensPerRequest: 4000,
    rateLimitPerMinute: 30,
  },
  defender: {
    type: "defender",
    name: "Defender Agent",
    scopes: [
      "https://graph.microsoft.com/SecurityEvents.Read.All",
      "https://graph.microsoft.com/Alert.Read.All",
      "https://graph.microsoft.com/Alert.ReadWrite.All",
      "https://graph.microsoft.com/ThreatIndicators.Read.All",
      "https://graph.microsoft.com/ThreatIndicators.ReadWrite.OwnedBy",
      "https://graph.microsoft.com/Vulnerability.Read.All",
      "https://graph.microsoft.com/Machine.Read.All",
      "https://graph.microsoft.com/Machine.ReadWrite.All",
    ],
    requiresApproval: true,
    maxTokensPerRequest: 4000,
    rateLimitPerMinute: 30,
  },
  m365admin: {
    type: "m365admin",
    name: "M365 Admin Agent",
    scopes: [
      "https://graph.microsoft.com/Organization.Read.All",
      "https://graph.microsoft.com/Organization.ReadWrite.All",
      "https://graph.microsoft.com/Domain.Read.All",
      "https://graph.microsoft.com/Domain.ReadWrite.All",
      "https://graph.microsoft.com/ServiceHealth.Read.All",
      "https://graph.microsoft.com/Reports.Read.All",
      "https://graph.microsoft.com/Directory.Read.All",
      "https://graph.microsoft.com/Directory.ReadWrite.All",
      "https://graph.microsoft.com/Policy.Read.All",
      "https://graph.microsoft.com/Policy.ReadWrite.All",
      "https://graph.microsoft.com/RoleManagement.Read.All",
      "https://graph.microsoft.com/RoleManagement.ReadWrite.Directory",
    ],
    requiresApproval: true,
    maxTokensPerRequest: 4000,
    rateLimitPerMinute: 30,
  },
  exchange: {
    type: "exchange",
    name: "Exchange Agent",
    scopes: [
      "https://graph.microsoft.com/Mail.Read",
      "https://graph.microsoft.com/Mail.ReadWrite",
      "https://graph.microsoft.com/Mail.Send",
      "https://graph.microsoft.com/MailboxSettings.Read",
      "https://graph.microsoft.com/MailboxSettings.ReadWrite",
      "https://graph.microsoft.com/Calendars.Read",
      "https://graph.microsoft.com/Calendars.ReadWrite",
      "https://graph.microsoft.com/Contacts.Read",
      "https://graph.microsoft.com/Contacts.ReadWrite",
      "https://graph.microsoft.com/Group.Read.All",
      "https://graph.microsoft.com/Group.ReadWrite.All",
    ],
    requiresApproval: true,
    maxTokensPerRequest: 4000,
    rateLimitPerMinute: 30,
  },
};

export interface AuditLogEntry {
  correlationId: string;
  userIdHash: string;
  tool: string;
  action: string;
  result: "success" | "error" | "blocked" | "preview";
  durationMs: number;
  timestamp: string;
  agentType: AgentType;
  metadata?: Record<string, unknown>;
}

export interface PreviewPlan {
  planId: string;
  hash: string;
  agentType: AgentType;
  action: string;
  target: string;
  payload: Record<string, unknown>;
  expiresAt: string;
  createdAt: string;
  status: "pending" | "approved" | "rejected" | "expired" | "executed";
  approvedBy?: string;
  approvedAt?: string;
}

export interface DLPMatch {
  type: "pem_key" | "jwt" | "connection_string" | "azure_secret" | "high_entropy" | "api_key" | "password";
  matchedText: string;
  position: { start: number; end: number };
  severity: "critical" | "high" | "medium";
}

export interface RateLimitState {
  tokens: number;
  lastRefill: number;
}

export interface CircuitBreakerState {
  failures: number;
  lastFailure: number;
  state: "closed" | "open" | "half-open";
  nextAttempt: number;
}