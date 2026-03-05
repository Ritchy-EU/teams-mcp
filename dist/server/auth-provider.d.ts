import type { Response } from "express";
import type { OAuthClientInformationFull, OAuthTokenRevocationRequest, OAuthTokens } from "@modelcontextprotocol/sdk/shared/auth.js";
import type { AuthorizationParams, OAuthServerProvider } from "@modelcontextprotocol/sdk/server/auth/provider.js";
import type { OAuthRegisteredClientsStore } from "@modelcontextprotocol/sdk/server/auth/clients.js";
import type { AuthInfo } from "@modelcontextprotocol/sdk/server/auth/types.js";
/**
 * OAuth provider that proxies authorization to Microsoft Entra ID.
 *
 * MCP clients dynamically register and get their own client_id, but all
 * requests to Microsoft use our Azure AD app's credentials (CLIENT_ID / AZURE_CLIENT_SECRET).
 *
 * Redirect flow:
 * 1. MCP client sends its own redirect_uri (e.g. http://localhost:40056/callback)
 * 2. We store it and redirect to Microsoft with OUR callback URL instead
 * 3. Microsoft redirects back to OUR callback URL with the code
 * 4. Our /oauth/callback handler redirects to the MCP client's original URL with the code
 * 5. MCP client exchanges the code via POST /oauth/token on our server
 * 6. We exchange the code with Microsoft using OUR callback URL (must match authorize)
 */
export declare class MicrosoftEntraOAuthProvider implements OAuthServerProvider {
    skipLocalPkceValidation: boolean;
    private _clientsStore;
    /**
     * Maps OAuth state → MCP client's original redirect_uri.
     * Entries are cleaned up after use or after TTL expiry.
     */
    private pendingAuthFlows;
    get clientsStore(): OAuthRegisteredClientsStore;
    authorize(_client: OAuthClientInformationFull, params: AuthorizationParams, res: Response): Promise<void>;
    /**
     * Called by the /oauth/callback handler to retrieve the MCP client's
     * original redirect_uri for a given state.
     *
     * Note: We intentionally do NOT delete the entry on first access.
     * Reverse proxies (ngrok, Cloudflare Tunnel) may show interstitial pages
     * that cause duplicate requests to /oauth/callback. The entry is cleaned
     * up by the TTL-based cleanup instead.
     */
    handleCallback(state: string): string | undefined;
    challengeForAuthorizationCode(_client: OAuthClientInformationFull, _authorizationCode: string): Promise<string>;
    exchangeAuthorizationCode(_client: OAuthClientInformationFull, authorizationCode: string, codeVerifier?: string, _redirectUri?: string): Promise<OAuthTokens>;
    exchangeRefreshToken(_client: OAuthClientInformationFull, refreshToken: string, scopes?: string[]): Promise<OAuthTokens>;
    verifyAccessToken(token: string): Promise<AuthInfo>;
    revokeToken?(_client: OAuthClientInformationFull, _request: OAuthTokenRevocationRequest): Promise<void>;
    private cleanupExpiredFlows;
}
//# sourceMappingURL=auth-provider.d.ts.map