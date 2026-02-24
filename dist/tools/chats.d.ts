import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";
/**
 * Registers all chat-related MCP tools on the given server.
 * Tools include: list_chats, get_chat_messages, send_chat_message,
 * create_chat, update_chat_message, and delete_chat_message.
 *
 * @param server - The MCP server instance to register tools on.
 * @param graphService - The Microsoft Graph service used for API calls.
 */
export declare function registerChatTools(server: McpServer, graphService: GraphService): void;
//# sourceMappingURL=chats.d.ts.map