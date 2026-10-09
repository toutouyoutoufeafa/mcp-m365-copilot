import { InteractiveBrowserCredential, TokenCredential, AuthenticationRecord } from "@azure/identity";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { CircuitBreaker } from "../security/circuitBreaker.js";
import { AgentType } from "../mcp/types.js";

export interface AuthConfig {
  clientId: string;
  tenantId: string;
  redirectUri?: string;
  enableWam: boolean;
}

export interface TokenResult {
  token: string;
  expiresOn: Date;
  scopes: string[];
  accountHash: string;
}

export class EntraAuthManager {
  private readonly credentials: Map<AgentType, TokenCredential> = new Map();
  private readonly authRecords: Map<AgentType, AuthenticationRecord> = new Map();
  private readonly tokenCache: Map<string, TokenResult> = new Map();
  private readonly circuitBreakers: Map<AgentType, CircuitBreaker> = new Map();
  private readonly config: AuthConfig;

  constructor(config: AuthConfig) {
    this.config = config;
    for (const agentType of ["dev", "knowledge", "collaboration", "purview", "autopilot", "defender", "m365admin", "exchange"] as AgentType[]) {
      this.circuitBreakers.set(agentType, new CircuitBreaker(3, 60000));
    }
  }

  async initialize(agentType: AgentType): Promise<void> {
    const correlationId = createCorrelationId();
    logger.info(`Initializing auth for agent: ${agentType}`, { correlationId, agentType });

    const scopes = this.getScopesForAgent(agentType);
    const credential = this.createCredential(agentType);
    this.credentials.set(agentType, credential);

    try {
      const token = await this.getTokenWithCircuitBreaker(agentType, scopes);
      logger.info(`Auth initialized successfully for ${agentType}`, {
        correlationId,
        agentType,
        expiresOn: token.expiresOn.toISOString(),
      });
    } catch (error) {
      logger.error(`Failed to initialize auth for ${agentType}`, {
        correlationId,
        agentType,
        error: error instanceof Error ? error.message : String(error),
      });
      throw error;
    }
  }

  async getToken(agentType: AgentType): Promise<string> {
    const scopes = this.getScopesForAgent(agentType);
    const tokenResult = await this.getTokenWithCircuitBreaker(agentType, scopes);
    return tokenResult.token;
  }

  async getTokenWithCircuitBreaker(agentType: AgentType, scopes: string[]): Promise<TokenResult> {
    const circuitBreaker = this.circuitBreakers.get(agentType)!;
    const cacheKey = `${agentType}:${scopes.sort().join(",")}`;

    const cached = this.tokenCache.get(cacheKey);
    if (cached && cached.expiresOn > new Date(Date.now() + 60000)) {
      return cached;
    }

    return circuitBreaker.execute(async () => {
      const credential = this.credentials.get(agentType);
      if (!credential) {
        throw new Error(`Credential not initialized for agent: ${agentType}`);
      }

      const correlationId = createCorrelationId();
      const startTime = Date.now();

      try {
        const tokenResponse = await credential.getToken(scopes);
        if (!tokenResponse) {
          throw new Error("Token response is null");
        }
        const accountHash = hashUserId(tokenResponse.token.slice(0, 50));

        const result: TokenResult = {
          token: tokenResponse.token,
          expiresOn: new Date(tokenResponse.expiresOnTimestamp),
          scopes,
          accountHash,
        };

        this.tokenCache.set(cacheKey, result);

        logger.audit({
          correlationId,
          userIdHash: accountHash,
          tool: "entra_auth",
          action: "get_token",
          result: "success",
          durationMs: Date.now() - startTime,
          agentType,
          metadata: { scopes, expiresOn: result.expiresOn.toISOString() },
        });

        return result;
      } catch (error) {
        logger.audit({
          correlationId,
          userIdHash: "unknown",
          tool: "entra_auth",
          action: "get_token",
          result: "error",
          durationMs: Date.now() - startTime,
          agentType,
          metadata: { error: error instanceof Error ? error.message : String(error) },
        });
        throw error;
      }
    });
  }

  private createCredential(agentType: AgentType): TokenCredential {
    const options: Record<string, unknown> = {
      clientId: this.config.clientId,
      tenantId: this.config.tenantId,
    };

    if (this.config.redirectUri) {
      options["redirectUri"] = this.config.redirectUri;
    }

    if (this.config.enableWam) {
      options["enableWam"] = true;
    }

    const record = this.authRecords.get(agentType);
    if (record) {
      options["authenticationRecord"] = record;
    }

    const credential = new InteractiveBrowserCredential(options as any);

    return credential;
  }

  private getScopesForAgent(agentType: AgentType): string[] {
    const scopeMap: Record<AgentType, string[]> = {
      dev: ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/Files.Read"],
      knowledge: ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/Files.Read"],
      collaboration: [
        "https://graph.microsoft.com/User.Read",
        "https://graph.microsoft.com/Files.Read",
        "https://graph.microsoft.com/Team.ReadBasic.All",
        "https://graph.microsoft.com/Channel.ReadBasic.All",
      ],
      purview: [
        "https://purview.azure.com/PurviewDataPolicies.Read.All",
        "https://purview.azure.com/PurviewCatalog.Read.All",
        "https://purview.azure.com/PurviewAccounts.Read.All",
      ],
      autopilot: [
        "https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All",
        "https://graph.microsoft.com/DeviceManagementManagedDevices.ReadWrite.All",
        "https://graph.microsoft.com/DeviceManagementServiceConfig.Read.All",
        "https://graph.microsoft.com/DeviceManagementServiceConfig.ReadWrite.All",
        "https://graph.microsoft.com/DeviceManagementConfiguration.Read.All",
        "https://graph.microsoft.com/DeviceManagementConfiguration.ReadWrite.All",
      ],
      defender: [
        "https://graph.microsoft.com/SecurityEvents.Read.All",
        "https://graph.microsoft.com/Alert.Read.All",
        "https://graph.microsoft.com/Alert.ReadWrite.All",
        "https://graph.microsoft.com/ThreatIndicators.Read.All",
        "https://graph.microsoft.com/ThreatIndicators.ReadWrite.OwnedBy",
        "https://graph.microsoft.com/Vulnerability.Read.All",
        "https://graph.microsoft.com/Machine.Read.All",
        "https://graph.microsoft.com/Machine.ReadWrite.All",
      ],
      m365admin: [
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
      exchange: [
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
    };
    return scopeMap[agentType];
  }

  async revokeTokens(agentType?: AgentType): Promise<void> {
    const correlationId = createCorrelationId();
    if (agentType) {
      this.tokenCache.clear();
      this.authRecords.delete(agentType);
      logger.info(`Revoked tokens for agent: ${agentType}`, { correlationId, agentType });
    } else {
      this.tokenCache.clear();
      this.authRecords.clear();
      logger.info("Revoked all tokens", { correlationId });
    }
  }

  async revokeInactiveTokens(maxInactiveMs: number = 300000): Promise<void> {
    const now = Date.now();
    for (const [key, token] of this.tokenCache.entries()) {
      if (now - token.expiresOn.getTime() > maxInactiveMs) {
        this.tokenCache.delete(key);
      }
    }
  }

  getCircuitBreakerState(agentType: AgentType): Readonly<{ state: string; failures: number }> {
    const cb = this.circuitBreakers.get(agentType);
    if (!cb) return { state: "unknown", failures: 0 };
    const state = cb.getState();
    return { state: state.state, failures: state.failures };
  }
}

export function createAuthManager(config: AuthConfig): EntraAuthManager {
  return new EntraAuthManager(config);
}