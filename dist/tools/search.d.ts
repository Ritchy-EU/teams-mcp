import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";
import type { SearchHit } from "../types/graph.js";
/**
 * Maps raw SearchHit objects from the Microsoft Search API into a
 * consistent, flat shape for tool responses.
 */
export declare function formatSearchHits(hits: SearchHit[]): {
    id: string;
    summary: string;
    rank: number;
    content: string | undefined;
    from: string | undefined;
    fromUserId: string | undefined;
    createdDateTime: string | undefined;
    importance: string | undefined;
    webLink: string | undefined;
    chatId: string | undefined;
    teamId: string | undefined;
    channelId: string | undefined;
}[];
export declare function registerSearchTools(server: McpServer, graphService: GraphService): void;
//# sourceMappingURL=search.d.ts.map