import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import type { SessionGraphService } from "../services/graph.js";

export interface SessionEntry {
  transport: StreamableHTTPServerTransport;
  server: McpServer;
  graphService: SessionGraphService;
  lastActivity: number;
  /** Update the access token used by this session's GraphService. */
  updateToken: (newToken: string) => void;
}

const SESSION_TIMEOUT_MS = 30 * 60 * 1000; // 30 minutes
const CLEANUP_INTERVAL_MS = 5 * 60 * 1000; // 5 minutes
const MAX_SESSIONS = 1000;

export class SessionManager {
  private sessions = new Map<string, SessionEntry>();
  private cleanupTimer: ReturnType<typeof setInterval> | undefined;

  constructor() {
    this.cleanupTimer = setInterval(() => this.cleanupStaleSessions(), CLEANUP_INTERVAL_MS);
  }

  get(sessionId: string): SessionEntry | undefined {
    const entry = this.sessions.get(sessionId);
    if (entry) {
      entry.lastActivity = Date.now();
    }
    return entry;
  }

  has(sessionId: string): boolean {
    return this.sessions.has(sessionId);
  }

  set(sessionId: string, entry: SessionEntry): void {
    if (this.sessions.size >= MAX_SESSIONS && !this.sessions.has(sessionId)) {
      throw new Error("Maximum session limit reached");
    }
    this.sessions.set(sessionId, entry);
  }

  delete(sessionId: string): void {
    this.sessions.delete(sessionId);
  }

  get size(): number {
    return this.sessions.size;
  }

  private cleanupStaleSessions(): void {
    const now = Date.now();
    for (const [sessionId, entry] of this.sessions) {
      if (now - entry.lastActivity > SESSION_TIMEOUT_MS) {
        console.log(`Cleaning up stale session: ${sessionId}`);
        entry.transport.close().catch((err) => {
          console.error(`Error closing stale transport for session ${sessionId}:`, err);
        });
        this.sessions.delete(sessionId);
      }
    }
  }

  async shutdown(): Promise<void> {
    if (this.cleanupTimer) {
      clearInterval(this.cleanupTimer);
      this.cleanupTimer = undefined;
    }

    const closePromises: Promise<void>[] = [];
    for (const [sessionId, entry] of this.sessions) {
      console.log(`Closing session: ${sessionId}`);
      closePromises.push(
        entry.transport.close().catch((err) => {
          console.error(`Error closing transport for session ${sessionId}:`, err);
        })
      );
    }

    await Promise.all(closePromises);
    this.sessions.clear();
  }
}
