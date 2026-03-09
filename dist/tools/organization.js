import { z } from "zod";
function toUserSummary(obj) {
    // DirectoryObject from the Graph API may contain User properties
    // when the underlying object is a User — cast via `any` to access them.
    const user = obj;
    return {
        id: user.id ?? undefined,
        displayName: user.displayName ?? undefined,
        userPrincipalName: user.userPrincipalName ?? undefined,
        mail: user.mail ?? undefined,
        jobTitle: user.jobTitle ?? undefined,
        department: user.department ?? undefined,
    };
}
export function registerOrganizationTools(server, graphService) {
    // Get user's manager
    server.tool("get_user_manager", "Get the manager of a user. Returns the user's direct manager with their profile information. If no userId is provided, returns the current user's manager.", {
        userId: z.string().optional().describe("User ID or email address. If omitted, uses the current authenticated user."),
    }, async ({ userId }) => {
        try {
            const client = await graphService.getClient();
            const endpoint = userId ? `/users/${userId}/manager` : "/me/manager";
            const manager = await client.api(endpoint).get();
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(toUserSummary(manager), null, 2),
                    },
                ],
            };
        }
        catch (error) {
            const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
            // Graph API returns 404 when user has no manager (top of org)
            if (errorMessage.includes("Resource") && errorMessage.includes("not found")) {
                return {
                    content: [
                        {
                            type: "text",
                            text: "No manager found. This user may be at the top of the organization hierarchy.",
                        },
                    ],
                };
            }
            return {
                content: [
                    {
                        type: "text",
                        text: `❌ Error: ${errorMessage}`,
                    },
                ],
            };
        }
    });
    // Get user's direct reports
    server.tool("get_direct_reports", "Get the direct reports of a user. Returns a list of users who report directly to the specified user. If no userId is provided, returns the current user's direct reports.", {
        userId: z.string().optional().describe("User ID or email address. If omitted, uses the current authenticated user."),
    }, async ({ userId }) => {
        try {
            const client = await graphService.getClient();
            const endpoint = userId ? `/users/${userId}/directReports` : "/me/directReports";
            const response = await client.api(endpoint).get();
            const directReports = response?.value ?? [];
            if (directReports.length === 0) {
                return {
                    content: [
                        {
                            type: "text",
                            text: "No direct reports found for this user.",
                        },
                    ],
                };
            }
            const reportsList = directReports.map((report) => toUserSummary(report));
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(reportsList, null, 2),
                    },
                ],
            };
        }
        catch (error) {
            const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
            return {
                content: [
                    {
                        type: "text",
                        text: `❌ Error: ${errorMessage}`,
                    },
                ],
            };
        }
    });
    // Get manager chain (upward hierarchy)
    server.tool("get_manager_chain", "Get the full chain of managers for a user, from their direct manager up to the top of the organization. Useful for understanding the reporting hierarchy.", {
        userId: z.string().optional().describe("User ID or email address. If omitted, uses the current authenticated user."),
        maxLevels: z.number().min(1).max(10).optional().describe("Maximum number of levels to traverse up the hierarchy (default: 5, max: 10)."),
    }, async ({ userId, maxLevels }) => {
        try {
            const client = await graphService.getClient();
            const limit = maxLevels ?? 5;
            const chain = [];
            // Start with the specified user or current user
            let currentEndpoint = userId ? `/users/${userId}/manager` : "/me/manager";
            for (let i = 0; i < limit; i++) {
                try {
                    const manager = await client.api(currentEndpoint).get();
                    const summary = toUserSummary(manager);
                    chain.push(summary);
                    // Next iteration: get this manager's manager
                    if (!manager.id)
                        break;
                    currentEndpoint = `/users/${manager.id}/manager`;
                }
                catch {
                    // No more managers in the chain (404 or similar)
                    break;
                }
            }
            if (chain.length === 0) {
                return {
                    content: [
                        {
                            type: "text",
                            text: "No managers found. This user may be at the top of the organization hierarchy.",
                        },
                    ],
                };
            }
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(chain, null, 2),
                    },
                ],
            };
        }
        catch (error) {
            const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
            return {
                content: [
                    {
                        type: "text",
                        text: `❌ Error: ${errorMessage}`,
                    },
                ],
            };
        }
    });
}
//# sourceMappingURL=organization.js.map