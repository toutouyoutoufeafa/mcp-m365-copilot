import { DLPMatch } from "../mcp/types.js";

const PEM_KEY_PATTERN = /-----BEGIN (?:RSA |EC |DSA |OPENSSH |PGP )?PRIVATE KEY-----[\s\S]*?-----END (?:RSA |EC |DSA |OPENSSH |PGP )?PRIVATE KEY-----/gi;
const JWT_PATTERN = /eyJ[A-Za-z0-9-_=]+\.[A-Za-z0-9-_=]+\.?[A-Za-z0-9-_.+/=]*/g;
const CONNECTION_STRING_PATTERN = /(?:Server|Data Source|Host|Endpoint|AccountName|AccountKey|SharedAccessKey|SecretKey|AccessKey|Password|ApiKey|Token)[\s:=]+[^\s;]{10,}/gi;
const AZURE_SECRET_PATTERN = /(?:azure|AZURE)[\s_-]?(?:secret|key|token|password|connection)[\s:=]+[A-Za-z0-9+/=]{20,}/gi;
const API_KEY_PATTERN = /(?:api[_-]?key|apikey|access[_-]?token|bearer[_-]?token)[\s:=]+[A-Za-z0-9-_]{20,}/gi;
const PASSWORD_PATTERN = /(?:password|passwd|pwd)[\s:=]+[^\s]{8,}/gi;
const HIGH_ENTROPY_PATTERN = /[A-Za-z0-9+/=]{32,}/g;

const PATTERNS: Array<{ regex: RegExp; type: DLPMatch["type"]; severity: DLPMatch["severity"] }> = [
  { regex: PEM_KEY_PATTERN, type: "pem_key", severity: "critical" },
  { regex: JWT_PATTERN, type: "jwt", severity: "critical" },
  { regex: CONNECTION_STRING_PATTERN, type: "connection_string", severity: "high" },
  { regex: AZURE_SECRET_PATTERN, type: "azure_secret", severity: "high" },
  { regex: API_KEY_PATTERN, type: "api_key", severity: "high" },
  { regex: PASSWORD_PATTERN, type: "password", severity: "high" },
];

function calculateEntropy(str: string): number {
  const freq = new Map<string, number>();
  for (const char of str) {
    freq.set(char, (freq.get(char) || 0) + 1);
  }
  let entropy = 0;
  const len = str.length;
  for (const count of freq.values()) {
    const p = count / len;
    entropy -= p * Math.log2(p);
  }
  return entropy;
}

export function scanForSecrets(text: string): DLPMatch[] {
  const matches: DLPMatch[] = [];

  for (const { regex, type, severity } of PATTERNS) {
    let match;
    while ((match = regex.exec(text)) !== null) {
      matches.push({
        type,
        matchedText: maskSecret(match[0]),
        position: { start: match.index, end: match.index + match[0].length },
        severity,
      });
    }
  }

  const entropyMatches = text.match(HIGH_ENTROPY_PATTERN);
  if (entropyMatches) {
    for (const match of entropyMatches) {
      const entropy = calculateEntropy(match);
      if (entropy > 4.5) {
        const index = text.indexOf(match);
        matches.push({
          type: "high_entropy",
          matchedText: maskSecret(match),
          position: { start: index, end: index + match.length },
          severity: entropy > 5.5 ? "critical" : "high",
        });
      }
    }
  }

  return matches;
}

function maskSecret(secret: string): string {
  if (secret.length <= 8) return "****";
  return secret.slice(0, 4) + "*".repeat(secret.length - 8) + secret.slice(-4);
}

export function sanitizeText(text: string): string {
  let sanitized = text;
  const matches = scanForSecrets(text);
  for (const match of matches) {
    const mask = "[REDACTED:" + match.type.toUpperCase() + "]";
    sanitized = sanitized.slice(0, match.position.start) + mask + sanitized.slice(match.position.end);
  }
  return sanitized;
}

export function hasCriticalSecrets(text: string): boolean {
  const matches = scanForSecrets(text);
  return matches.some((m) => m.severity === "critical");
}