export declare const CLIENT_ID: string;
export declare const TENANT_ID: string;
export declare const AUTHORITY: string;
/** Scopes sufficient for read-only operations (no sending, no uploads). */
export declare const READ_ONLY_SCOPES: string[];
/** Full scopes including write operations. */
export declare const FULL_SCOPES: string[];
/**
 * Read-only mode: set TEAMS_MCP_READ_ONLY=true (or 1/yes) to skip registering
 * write tools and request only the reduced permission scopes.
 */
export declare const READ_ONLY: boolean;
export declare const DELEGATED_SCOPES: string[];
/**
 * Scopes requested from Microsoft Entra ID in HTTP (OAuth) mode.
 *
 * `offline_access` is what makes Entra issue a refresh token; without it the
 * session dies unrecoverably when the access token expires (~1h), so it is
 * always forced into the outbound request regardless of what the MCP client asked for.
 */
export declare const HTTP_SCOPES: string[];
/**
 * Directory for server-side state that must survive a restart (currently the OAuth
 * client registry). In Docker this should be a mounted volume — see docker-compose.yml.
 */
export declare const DATA_DIR: string;
export declare const PORT: number;
export declare const BASE_URL: string;
export declare const AZURE_CLIENT_SECRET: string | undefined;
//# sourceMappingURL=config.d.ts.map