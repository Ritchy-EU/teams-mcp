import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";

export function registerAuthTools(server: McpServer, graphService: GraphService) {
  // Authentication status tool
  server.tool(
    "auth_status",
    "Check the authentication status of the Microsoft Graph connection. Returns whether the user is authenticated and shows their basic profile information.",
    {},
    async () => {
      const status = await graphService.getAuthStatus();
      return {
        content: [
          {
            type: "text",
            text: status.isAuthenticated
              ? `✅ Authenticated as ${status.displayName || "Unknown User"} (${status.userPrincipalName || "No email available"})`
              : "❌ Not authenticated. Use the start_authentication tool or run /ms-teams:authenticate to sign in.",
          },
        ],
      };
    }
  );

  // Start authentication tool
  server.tool(
    "start_authentication",
    "Start Microsoft Teams authentication via device code flow. Returns a URL and code for the user to complete sign-in in their browser. Use this when auth_status shows not authenticated.",
    {},
    async () => {
      try {
        const { verificationUri, userCode, expiresIn } = await graphService.startDeviceCodeAuth();
        return {
          content: [
            {
              type: "text",
              text: [
                "🔐 Authentication required. Please complete these steps:",
                "",
                `1. Open this URL: ${verificationUri}`,
                `2. Enter this code: ${userCode}`,
                "",
                `The code expires in ${Math.floor(expiresIn / 60)} minutes.`,
                "Once you complete sign-in in the browser, authentication will be saved automatically.",
              ].join("\n"),
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `❌ Failed to start authentication: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
        };
      }
    }
  );
}
