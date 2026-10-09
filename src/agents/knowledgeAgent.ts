import { MCPTool, ToolResult } from "../mcp/types.js";
import { BaseAgent } from "./baseAgent.js";
import { EntraAuthManager } from "../auth/entraAuth.js";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { sanitizeText } from "../dlp/secretScanner.js";
import { readFile as fsReadFile, readdir, stat as fsStat } from "fs/promises";
import { join, resolve, relative, extname } from "path";

const ALLOWED_DIRS = ["Standards", "Runbooks", "Templates", "Architecture"];
const ALLOWED_EXTENSIONS = [".md", ".txt", ".json", ".yaml", ".yml", ".pdf", ".docx"];
const MAX_FILE_SIZE = 5 * 1024 * 1024; // 5MB
const MAX_RESULTS = 50;

export class KnowledgeAgent extends BaseAgent {
  private readonly workspaceRoot: string;

  constructor(authManager: EntraAuthManager, workspaceRoot: string) {
    const config = {
      type: "knowledge" as const,
      name: "Knowledge Agent",
      scopes: ["User.Read", "Files.Read"],
      allowlist: ["Standards/", "Runbooks/", "Templates/", "Architecture/"],
      requiresApproval: false,
      maxTokensPerRequest: 4000,
      rateLimitPerMinute: 60,
    };
    super(config, authManager, "knowledge");
    this.workspaceRoot = resolve(workspaceRoot);
  }

  override getTools(): MCPTool[] {
    return [
      {
        name: "engineering_knowledge_search",
        description: "Search allowlisted knowledge directories (Standards, Runbooks, Templates, Architecture) for relevant content. Does NOT call Copilot Chat or /search/query.",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Search query (keywords, phrases)" },
            directories: {
              type: "array",
              items: { type: "string", enum: ALLOWED_DIRS },
              description: "Directories to search (default: all allowlisted)",
            },
            maxResults: { type: "number", description: "Maximum results to return", default: 20, maximum: MAX_RESULTS },
            includeContent: { type: "boolean", description: "Include file content in results", default: true },
          },
          required: ["query"],
        },
      },
      {
        name: "engineering_knowledge_read",
        description: "Read a specific file from allowlisted knowledge directories",
        inputSchema: {
          type: "object",
          properties: {
            path: { type: "string", description: "Relative path from workspace root (must be in allowlist)" },
          },
          required: ["path"],
        },
      },
      {
        name: "engineering_knowledge_list",
        description: "List all files in allowlisted knowledge directories",
        inputSchema: {
          type: "object",
          properties: {
            directory: { type: "string", enum: ALLOWED_DIRS, description: "Specific directory to list" },
          },
        },
      },
    ];
  }

  async executeTool(name: string, args: Record<string, unknown>): Promise<ToolResult> {
    const getArg = (key: string) => args[key];

    switch (name) {
      case "engineering_knowledge_search":
        return this.executeWithGuardrails(name, args, () =>
          this.searchKnowledge(getArg("query") as string, getArg("directories") as string[] | undefined, (getArg("maxResults") as number) || 20, (getArg("includeContent") as boolean) ?? true)
        );
      case "engineering_knowledge_read":
        return this.executeWithGuardrails(name, args, () => this.readKnowledgeFile(getArg("path") as string));
      case "engineering_knowledge_list":
        return this.executeWithGuardrails(name, args, () => this.listKnowledge(getArg("directory") as string));
      default:
        return {
          success: false,
          error: `Unknown tool: ${name}`,
          auditEntry: {
            correlationId: createCorrelationId(),
            userIdHash: hashUserId(createCorrelationId()),
            tool: `knowledge.${name}`,
            action: name,
            result: "error",
            durationMs: 0,
            agentType: "knowledge",
          },
        };
    }
  }

  private async searchKnowledge(
    query: string,
    directories?: string[],
    maxResults: number = 20,
    includeContent: boolean = true
  ): Promise<{ results: Array<{ path: string; score: number; snippet?: string }>; total: number }> {
    const searchDirs = directories && directories.length > 0 ? directories : ALLOWED_DIRS;
    const queryTerms = query.toLowerCase().split(/\s+/).filter((t) => t.length > 2);
    const results: Array<{ path: string; score: number; snippet?: string }> = [];

    for (const dir of searchDirs) {
      const dirPath = join(this.workspaceRoot, dir);
      const files = await this.collectFiles(dirPath);
      for (const file of files) {
        const score = await this.scoreFile(file, queryTerms);
        if (score > 0) {
          const relPath = relative(this.workspaceRoot, file);
          let snippet: string | undefined;
          if (includeContent) {
            const content = await this.readFileSafe(file);
            snippet = this.extractSnippet(content, queryTerms);
          }
          results.push({ path: relPath, score, snippet });
        }
      }
    }

    results.sort((a, b) => b.score - a.score);
    return { results: results.slice(0, maxResults), total: results.length };
  }

  private async collectFiles(dir: string): Promise<string[]> {
    const files: string[] = [];
    try {
      const entries = await readdir(dir, { withFileTypes: true });
      for (const entry of entries) {
        const fullPath = join(dir, entry.name);
        if (entry.isDirectory()) {
          files.push(...(await this.collectFiles(fullPath)));
        } else if (entry.isFile()) {
          const ext = extname(entry.name).toLowerCase();
          if (ALLOWED_EXTENSIONS.includes(ext)) {
            const fileStat = await fsStat(fullPath);
            if (fileStat.size <= MAX_FILE_SIZE) {
              files.push(fullPath);
            }
          }
        }
      }
    } catch {
      // Directory doesn't exist, skip
    }
    return files;
  }

  private async scoreFile(file: string, queryTerms: string[]): Promise<number> {
    const content = await this.readFileSafe(file);
    if (!content) return 0;

    const lowerContent = content.toLowerCase();
    let score = 0;
    for (const term of queryTerms) {
      const matches = (lowerContent.match(new RegExp(term, "g")) || []).length;
      score += matches;
    }
    return score;
  }

  private async readFileSafe(file: string): Promise<string> {
    try {
      return await fsReadFile(file, "utf-8");
    } catch {
      return "";
    }
  }

  private extractSnippet(content: string, queryTerms: string[]): string {
    const lines = content.split("\n");
    if (lines.length === 0) return "";
    let bestLine = 0;
    let bestScore = 0;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i]?.toLowerCase() ?? "";
      let score = 0;
      for (const term of queryTerms) {
        if (line.includes(term)) score++;
      }
      if (score > bestScore) {
        bestScore = score;
        bestLine = i;
      }
    }

    const start = Math.max(0, bestLine - 2);
    const end = Math.min(lines.length, bestLine + 3);
    return sanitizeText(lines.slice(start, end).join("\n"));
  }

  private async readKnowledgeFile(path: string): Promise<string> {
    const fullPath = this.resolveAndValidatePath(path);
    const fileStat = await fsStat(fullPath);
    if (fileStat.size > MAX_FILE_SIZE) {
      throw new Error(`File too large: ${fileStat.size} bytes`);
    }
    const content = await fsReadFile(fullPath, "utf-8");
    return sanitizeText(content);
  }

  private async listKnowledge(directory?: string): Promise<{ path: string; name: string; size: number; modified: string }[]> {
    const dirs = directory ? [directory] : ALLOWED_DIRS;
    const results: { path: string; name: string; size: number; modified: string }[] = [];

    for (const dir of dirs) {
      const dirPath = join(this.workspaceRoot, dir);
      const files = await this.collectFiles(dirPath);
      for (const file of files) {
        const fileStat = await fsStat(file);
        results.push({
          path: relative(this.workspaceRoot, file),
          name: file.split("/").pop() || "",
          size: fileStat.size,
          modified: fileStat.mtime.toISOString(),
        });
      }
    }

    return results;
  }

  private resolveAndValidatePath(path: string): string {
    const normalized = path.replace(/\\/g, "/");
    const fullPath = resolve(this.workspaceRoot, normalized);

    if (!fullPath.startsWith(this.workspaceRoot)) {
      throw new Error(`Path traversal attempt blocked: ${path}`);
    }

    const relPath = relative(this.workspaceRoot, fullPath);
    const allowed = ALLOWED_DIRS.some((d) => relPath.startsWith(d));
    if (!allowed) {
      throw new Error(`Path not in allowlisted directories: ${relPath}`);
    }

    return fullPath;
  }
}