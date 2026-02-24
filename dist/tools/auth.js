export function registerAuthTools(server, graphService) {
    // Authentication status tool
    server.tool("auth_status", "Check the authentication status of the Microsoft Graph connection. Returns whether the user is authenticated and shows their basic profile information.", {}, async () => {
        const status = await graphService.getAuthStatus();
        return {
            content: [
                {
                    type: "text",
                    text: status.isAuthenticated
                        ? `✅ Authenticated as ${status.displayName || "Unknown User"} (${status.userPrincipalName || "No email available"})`
                        : "❌ Not authenticated. Please run: npx @ritchy-eu/teams-mcp@latest authenticate",
                },
            ],
        };
    });
}
//# sourceMappingURL=auth.js.map