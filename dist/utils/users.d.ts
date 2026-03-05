import type { IGraphService } from "../services/graph.js";
export interface UserInfo {
    id: string;
    displayName: string;
    userPrincipalName?: string;
}
/**
 * Search for users by display name or email
 */
export declare function searchUsers(graphService: IGraphService, query: string, limit?: number): Promise<UserInfo[]>;
/**
 * Get user by exact email or UPN
 */
export declare function getUserByEmail(graphService: IGraphService, email: string): Promise<UserInfo | null>;
/**
 * Get user by ID
 */
export declare function getUserById(graphService: IGraphService, userId: string): Promise<UserInfo | null>;
/**
 * Parse @mentions from text and return user lookup suggestions
 * @param text - Message text containing @mentions
 * @param graphService - Graph service instance
 * @returns Array of mention patterns found and suggested users
 */
export declare function parseMentions(text: string, graphService: IGraphService): Promise<Array<{
    mention: string;
    users: UserInfo[];
}>>;
/**
 * Generate HTML content with @mentions converted to proper format
 */
export declare function processMentionsInHtml(html: string, mentionMappings: Array<{
    mention: string;
    userId: string;
    displayName: string;
}>): {
    content: string;
    mentions: Array<{
        id: number;
        mentionText: string;
        mentioned: {
            user: {
                id: string;
            };
        };
    }>;
};
//# sourceMappingURL=users.d.ts.map