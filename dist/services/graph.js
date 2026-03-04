import { PublicClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { AUTHORITY, CLIENT_ID } from "../config.js";
import { cachePlugin } from "../msal-cache.js";
const DELEGATED_SCOPES = [
    "User.Read",
    "User.ReadBasic.All",
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "ChannelMessage.Read.All",
    "TeamMember.Read.All",
    "Chat.ReadBasic",
    "Chat.ReadWrite",
];
export class GraphService {
    static instance;
    client;
    isInitialized = false;
    tokenExpiresAt;
    msalApp;
    msalAccount;
    pendingDeviceCodeAuth;
    static getInstance() {
        if (!GraphService.instance) {
            GraphService.instance = new GraphService();
        }
        return GraphService.instance;
    }
    async initializeClient() {
        if (this.isInitialized)
            return;
        try {
            // Priority 1: AUTH_TOKEN environment variable (direct token injection)
            const envToken = process.env.AUTH_TOKEN;
            if (envToken) {
                const validatedToken = this.validateToken(envToken);
                if (validatedToken) {
                    this.client = Client.initWithMiddleware({
                        authProvider: {
                            getAccessToken: async () => validatedToken,
                        },
                    });
                    this.isInitialized = true;
                }
                return;
            }
            // Priority 2: MSAL with cached refresh token for automatic token renewal
            this.msalApp = new PublicClientApplication({
                auth: {
                    clientId: CLIENT_ID,
                    authority: AUTHORITY,
                },
                cache: {
                    cachePlugin,
                },
            });
            const accounts = await this.msalApp.getTokenCache().getAllAccounts();
            if (accounts.length === 0) {
                return;
            }
            this.msalAccount = accounts[0];
            // Verify we can acquire a token
            const result = await this.msalApp.acquireTokenSilent({
                scopes: DELEGATED_SCOPES,
                account: this.msalAccount,
            });
            if (!result) {
                return;
            }
            this.tokenExpiresAt = result.expiresOn ?? undefined;
            // Create Graph client with MSAL-backed auth provider for automatic token refresh
            this.client = Client.initWithMiddleware({
                authProvider: {
                    getAccessToken: () => this.acquireToken(),
                },
            });
            this.isInitialized = true;
        }
        catch (error) {
            console.error("Failed to initialize Graph client:", error);
        }
    }
    async acquireToken() {
        if (!this.msalApp || !this.msalAccount) {
            throw new Error("MSAL not initialized");
        }
        const result = await this.msalApp.acquireTokenSilent({
            scopes: DELEGATED_SCOPES,
            account: this.msalAccount,
        });
        if (!result) {
            throw new Error("Failed to acquire access token. Run /ms-teams:authenticate in Claude Code to sign in again.");
        }
        this.tokenExpiresAt = result.expiresOn ?? undefined;
        return result.accessToken;
    }
    async getAuthStatus() {
        await this.initializeClient();
        if (!this.client) {
            return { isAuthenticated: false };
        }
        try {
            const me = await this.client.api("/me").get();
            return {
                isAuthenticated: true,
                userPrincipalName: me?.userPrincipalName ?? undefined,
                displayName: me?.displayName ?? undefined,
                expiresAt: this.tokenExpiresAt?.toISOString(),
            };
        }
        catch (error) {
            console.error("Error getting user info:", error);
            return { isAuthenticated: false };
        }
    }
    async getClient() {
        await this.initializeClient();
        if (!this.client) {
            throw new Error("Not authenticated. Run /ms-teams:authenticate in Claude Code to sign in.");
        }
        return this.client;
    }
    isAuthenticated() {
        return !!this.client && this.isInitialized;
    }
    async startDeviceCodeAuth() {
        return new Promise((resolve, reject) => {
            const msalApp = new PublicClientApplication({
                auth: {
                    clientId: CLIENT_ID,
                    authority: AUTHORITY,
                },
                cache: {
                    cachePlugin,
                },
            });
            this.pendingDeviceCodeAuth = msalApp.acquireTokenByDeviceCode({
                scopes: DELEGATED_SCOPES,
                deviceCodeCallback: (response) => {
                    resolve({
                        verificationUri: response.verificationUri,
                        userCode: response.userCode,
                        expiresIn: response.expiresIn,
                    });
                },
            });
            // If device code request itself fails (network error etc.), reject immediately
            this.pendingDeviceCodeAuth.catch(reject);
            // When auth completes successfully, re-initialize the Graph client
            this.pendingDeviceCodeAuth
                .then(async (result) => {
                if (result) {
                    this.isInitialized = false;
                    await this.initializeClient();
                }
            })
                .catch((err) => {
                console.error("Device code authentication failed:", err);
            });
        });
    }
    validateToken(token) {
        const tokenSplits = token.split(".");
        if (tokenSplits.length !== 3) {
            console.error("Invalid JWT token: missing claims");
            return undefined;
        }
        try {
            const payload = JSON.parse(atob(tokenSplits[1]));
            // Check audience
            const audiences = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
            if (!audiences.includes("https://graph.microsoft.com")) {
                console.error("Invalid JWT token: Not a valid Microsoft Graph token");
                return undefined;
            }
            // Check expiration
            if (typeof payload.exp === "number" && payload.exp * 1000 < Date.now()) {
                console.error("Invalid JWT token: Token has expired");
                return undefined;
            }
            // Check issuer
            if (typeof payload.iss === "string") {
                const validIssuers = ["https://login.microsoftonline.com/", "https://sts.windows.net/"];
                if (!validIssuers.some((prefix) => payload.iss.startsWith(prefix))) {
                    console.error("Invalid JWT token: Unrecognized issuer");
                    return undefined;
                }
            }
        }
        catch (error) {
            console.error("Invalid JWT token: Failed to parse payload", error);
            return undefined;
        }
        return token;
    }
}
//# sourceMappingURL=graph.js.map