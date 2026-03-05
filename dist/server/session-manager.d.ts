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
export declare class SessionManager {
    private sessions;
    private cleanupTimer;
    constructor();
    get(sessionId: string): SessionEntry | undefined;
    has(sessionId: string): boolean;
    set(sessionId: string, entry: SessionEntry): void;
    delete(sessionId: string): void;
    get size(): number;
    private cleanupStaleSessions;
    shutdown(): Promise<void>;
}
//# sourceMappingURL=session-manager.d.ts.map