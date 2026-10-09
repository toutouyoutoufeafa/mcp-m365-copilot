import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  InitializeRequestSchema,
  InitializeResult,
  Tool,
} from "@modelcontextprotocol/sdk/types.js";
import { EntraAuthManager, createAuthManager, AuthConfig } from "../auth/entraAuth.js";
import { DevAgent } from "../agents/devAgent.js";
import { KnowledgeAgent } from "../agents/knowledgeAgent.js";
import { CollaborationAgent } from "../agents/collaborationAgent.js";
import { PurviewAgent } from "../agents/purviewAgent.js";
import { AutopilotAgent } from "../agents/autopilotAgent.js";
import { DefenderAgent } from "../agents/defenderAgent.js";
import { M365AdminAgent } from "../agents/m365adminAgent.js";
import { ExchangeAgent } from "../agents/exchangeAgent.js";
import { logger, createCorrelationId } from "../utils/logger.js";
import { MCPTool, AgentType } from "./types.js";

export class MCPServer {
  private readonly server: Server;
  private readonly authManager: EntraAuthManager;
  private readonly devAgent: DevAgent;
  private readonly knowledgeAgent: KnowledgeAgent;
  private readonly collaborationAgent: CollaborationAgent;
  private readonly purviewAgent: PurviewAgent | null;
  private readonly autopilotAgent: AutopilotAgent | null;
  private readonly defenderAgent: DefenderAgent | null;
  private readonly m365adminAgent: M365AdminAgent | null;
  private readonly exchangeAgent: ExchangeAgent | null;
  private readonly agents: Map<AgentType, DevAgent | KnowledgeAgent | CollaborationAgent | PurviewAgent | AutopilotAgent | DefenderAgent | M365AdminAgent | ExchangeAgent>;
  private readonly workspaceRoot: string;
  private initialized = false;

  constructor(workspaceRoot: string, authConfig: AuthConfig) {
    this.workspaceRoot = workspaceRoot;
    this.authManager = createAuthManager(authConfig);

    this.devAgent = new DevAgent(this.authManager, workspaceRoot);
    this.knowledgeAgent = new KnowledgeAgent(this.authManager, workspaceRoot);
    this.collaborationAgent = new CollaborationAgent(this.authManager);

    // Optional agents - enabled via environment variables
    const purviewAccountName = process.env["PURVIEW_ACCOUNT_NAME"];
    this.purviewAgent = purviewAccountName 
      ? new PurviewAgent(this.authManager, purviewAccountName)
      : null;

    const enableAutopilot = process.env["ENABLE_AUTOPILOT"] === "true";
    this.autopilotAgent = enableAutopilot 
      ? new AutopilotAgent(this.authManager)
      : null;

    const enableDefender = process.env["ENABLE_DEFENDER"] === "true";
    this.defenderAgent = enableDefender 
      ? new DefenderAgent(this.authManager)
      : null;

    const enableM365Admin = process.env["ENABLE_M365ADMIN"] === "true";
    this.m365adminAgent = enableM365Admin 
      ? new M365AdminAgent(this.authManager)
      : null;

    const enableExchange = process.env["ENABLE_EXCHANGE"] === "true";
    this.exchangeAgent = enableExchange 
      ? new ExchangeAgent(this.authManager)
      : null;

    const agentMap = new Map<AgentType, DevAgent | KnowledgeAgent | CollaborationAgent | PurviewAgent | AutopilotAgent | DefenderAgent | M365AdminAgent | ExchangeAgent>([
      ["dev", this.devAgent],
      ["knowledge", this.knowledgeAgent],
      ["collaboration", this.collaborationAgent],
    ]);

    if (this.purviewAgent) {
      agentMap.set("purview", this.purviewAgent);
      logger.info("Purview Agent enabled", { accountName: purviewAccountName });
    }
    if (this.autopilotAgent) {
      agentMap.set("autopilot", this.autopilotAgent);
      logger.info("Autopilot Agent enabled");
    }
    if (this.defenderAgent) {
      agentMap.set("defender", this.defenderAgent);
      logger.info("Defender Agent enabled");
    }
    if (this.m365adminAgent) {
      agentMap.set("m365admin", this.m365adminAgent);
      logger.info("M365 Admin Agent enabled");
    }
    if (this.exchangeAgent) {
      agentMap.set("exchange", this.exchangeAgent);
      logger.info("Exchange Agent enabled");
    }

    this.agents = agentMap;

    this.server = new Server(
      {
        name: "mcp-m365-copilot",
        version: "1.0.0",
      },
      {
        capabilities: {
          tools: { listChanged: true },
        },
      }
    );

    this.setupHandlers();
  }

  private setupHandlers(): void {
    this.server.setRequestHandler(InitializeRequestSchema, async (request) => {
      const correlationId = createCorrelationId();
      logger.info("MCP Initialize request received", { correlationId, clientInfo: request.params.clientInfo });

      for (const agent of this.agents.values()) {
        await agent.initialize();
      }

      this.initialized = true;

      const result: InitializeResult = {
        protocolVersion: "2024-11-05",
        capabilities: {
          tools: { listChanged: true },
        },
        serverInfo: {
          name: "mcp-m365-copilot",
          version: "1.0.0",
        },
      };

      logger.info("MCP Initialize completed", { correlationId });
      return result;
    });

    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      const correlationId = createCorrelationId();
      logger.debug("Listing tools", { correlationId });

      const allTools: Tool[] = [];
      for (const agent of this.agents.values()) {
        const tools = agent.getTools();
        allTools.push(...this.convertTools(tools));
      }

      return { tools: allTools };
    });

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const correlationId = createCorrelationId();
      const { name, arguments: args } = request.params;

      logger.info("Tool call received", { correlationId, tool: name });

      const agent = this.getAgentForTool(name);
      if (!agent) {
        throw new Error(`No agent found for tool: ${name}`);
      }

      const result = await agent.executeTool(name, args || {});

      if (!result.success) {
        throw new Error(result.error || "Tool execution failed");
      }

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result.data, null, 2),
          },
        ],
      };
    });
  }

  private convertTools(tools: MCPTool[]): Tool[] {
    return tools.map((t) => ({
      name: t.name,
      description: t.description,
      inputSchema: t.inputSchema as Tool["inputSchema"],
    }));
  }

  private getAgentForTool(name: string): DevAgent | KnowledgeAgent | CollaborationAgent | PurviewAgent | AutopilotAgent | DefenderAgent | M365AdminAgent | ExchangeAgent | null {
    if (name.startsWith("dev_")) return this.devAgent;
    if (name.startsWith("engineering_knowledge_")) return this.knowledgeAgent;
    if (name.startsWith("collab_")) return this.collaborationAgent;
    if (name.startsWith("purview_")) return this.purviewAgent;
    if (name.startsWith("autopilot_")) return this.autopilotAgent;
    if (name.startsWith("defender_")) return this.defenderAgent;
    if (name.startsWith("m365admin_")) return this.m365adminAgent;
    if (name.startsWith("exchange_")) return this.exchangeAgent;
    return null;
  }

  async start(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    logger.info("MCP Server started on stdio");
  }

  async stop(): Promise<void> {
    await this.authManager.revokeTokens();
    logger.info("MCP Server stopped");
  }
}

export async function createMCPServer(workspaceRoot: string, authConfig: AuthConfig): Promise<MCPServer> {
  return new MCPServer(workspaceRoot, authConfig);
}