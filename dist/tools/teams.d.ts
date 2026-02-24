import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";
/**
 * Registers all Teams-related MCP tools on the given server.
 * Tools include: list_teams, list_channels, get_channel_messages,
 * send_channel_message, get_channel_message_replies, reply_to_channel_message,
 * list_team_members, search_users_for_mentions, download_message_hosted_content,
 * delete_channel_message, and update_channel_message.
 *
 * @param server - The MCP server instance to register tools on.
 * @param graphService - The Microsoft Graph service used for API calls.
 */
export declare function registerTeamsTools(server: McpServer, graphService: GraphService): void;
//# sourceMappingURL=teams.d.ts.map