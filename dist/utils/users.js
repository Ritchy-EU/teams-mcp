/**
 * Escape a string value for safe use inside OData single-quoted literals.
 * OData escapes single quotes by doubling them: ' → ''
 */
function escapeODataString(value) {
    return value.replace(/'/g, "''");
}
/**
 * Search for users by display name or email
 */
export async function searchUsers(graphService, query, limit = 10) {
    try {
        const client = await graphService.getClient();
        // Use filter query to search users by displayName or userPrincipalName
        const safeQuery = escapeODataString(query);
        const searchQuery = `$filter=startswith(displayName,'${safeQuery}') or startswith(userPrincipalName,'${safeQuery}')&$top=${limit}&$select=id,displayName,userPrincipalName`;
        const response = await client.api(`/users?${searchQuery}`).get();
        if (!response?.value?.length) {
            return [];
        }
        return response.value.map((user) => ({
            id: user.id || "",
            displayName: user.displayName || "Unknown User",
            userPrincipalName: user.userPrincipalName || undefined,
        }));
    }
    catch (error) {
        console.error("Error searching users:", error);
        return [];
    }
}
/**
 * Get user by exact email or UPN
 */
export async function getUserByEmail(graphService, email) {
    try {
        const client = await graphService.getClient();
        const response = await client.api(`/users/${email}`).get();
        return {
            id: response.id,
            displayName: response.displayName || "Unknown User",
            userPrincipalName: response.userPrincipalName,
        };
    }
    catch (_error) {
        // User not found or access denied
        return null;
    }
}
/**
 * Get user by ID
 */
export async function getUserById(graphService, userId) {
    try {
        const client = await graphService.getClient();
        const response = await client
            .api(`/users/${userId}`)
            .select("id,displayName,userPrincipalName")
            .get();
        return {
            id: response.id,
            displayName: response.displayName || "Unknown User",
            userPrincipalName: response.userPrincipalName,
        };
    }
    catch (_error) {
        // User not found or access denied
        return null;
    }
}
/**
 * Parse @mentions from text and return user lookup suggestions
 * @param text - Message text containing @mentions
 * @param graphService - Graph service instance
 * @returns Array of mention patterns found and suggested users
 */
export async function parseMentions(text, graphService) {
    // Match @mentions in the format @username, @email@domain.com, or @"User Name"
    const mentionRegex = /@(?:"([^"]+)"|([^\s@]+(?:@[^\s@]+\.[^\s@]+)?|[^\s@]+))/g;
    const mentions = [];
    let match = mentionRegex.exec(text);
    while (match !== null) {
        const mentionText = match[1] || match[2]; // Quoted name or unquoted
        let users = [];
        // If it looks like an email, try exact lookup first
        if (mentionText.includes("@") && mentionText.includes(".")) {
            const user = await getUserByEmail(graphService, mentionText);
            if (user) {
                users = [user];
            }
        }
        // If no exact match found, search by name
        if (users.length === 0) {
            users = await searchUsers(graphService, mentionText, 5);
        }
        mentions.push({
            mention: mentionText,
            users,
        });
        match = mentionRegex.exec(text);
    }
    return mentions;
}
/**
 * Generate HTML content with @mentions converted to proper format
 */
export function processMentionsInHtml(html, mentionMappings) {
    let processedContent = html;
    const mentions = [];
    mentionMappings.forEach((mapping, index) => {
        // Replace @mention with HTML mention format
        const mentionRegex = new RegExp(`@(?:"${escapeRegex(mapping.mention)}"|${escapeRegex(mapping.mention)})`, "g");
        const mentionId = index;
        processedContent = processedContent.replace(mentionRegex, `<at id="${mentionId}">${mapping.displayName}</at>`);
        mentions.push({
            id: mentionId,
            mentionText: mapping.displayName,
            mentioned: {
                user: {
                    id: mapping.userId,
                },
            },
        });
    });
    return { content: processedContent, mentions };
}
function escapeRegex(text) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
//# sourceMappingURL=users.js.map