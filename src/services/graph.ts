import { type AccountInfo, PublicClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { AUTHORITY, CLIENT_ID, DELEGATED_SCOPES } from "../config.js";
import { cachePlugin } from "../msal-cache.js";

export interface AuthStatus {
  isAuthenticated: boolean;
  userPrincipalName?: string | undefined;
  displayName?: string | undefined;
  expiresAt?: string | undefined;
}

/**
 * Common interface for Graph API access, used by both stdio (singleton) and HTTP (per-session) modes.
 */
export interface IGraphService {
  getAuthStatus(): Promise<AuthStatus>;
  getClient(): Promise<Client>;
  isAuthenticated(): boolean;
  startDeviceCodeAuth(): Promise<{
    verificationUri: string;
    userCode: string;
    expiresIn: number;
  }>;
  validateToken(token: string): string | undefined;
}

/**
 * Validate a Microsoft Graph JWT token (shared logic).
 */
export function validateGraphToken(token: string): string | undefined {
  const tokenSplits = token.split(".");
  if (tokenSplits.length !== 3) {
    console.error("Invalid JWT token: missing claims");
    return undefined;
  }

  try {
    const payload = JSON.parse(atob(tokenSplits[1]));

    // Microsoft Graph tokens may use either the URL or the GUID as audience:
    //   v1 format: "https://graph.microsoft.com"
    //   v2 format: "00000003-0000-0000-c000-000000000000" (Graph API app ID)
    const VALID_GRAPH_AUDIENCES = [
      "https://graph.microsoft.com",
      "00000003-0000-0000-c000-000000000000",
    ];
    const audiences = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
    if (!audiences.some((aud: string) => VALID_GRAPH_AUDIENCES.includes(aud))) {
      console.error(`Invalid JWT token: Not a valid Microsoft Graph token (aud=${JSON.stringify(payload.aud)})`);
      return undefined;
    }

    if (typeof payload.exp === "number" && payload.exp * 1000 < Date.now()) {
      console.error("Invalid JWT token: Token has expired");
      return undefined;
    }

    if (typeof payload.iss === "string") {
      const validIssuers = ["https://login.microsoftonline.com/", "https://sts.windows.net/"];
      if (!validIssuers.some((prefix) => payload.iss.startsWith(prefix))) {
        console.error("Invalid JWT token: Unrecognized issuer");
        return undefined;
      }
    }
  } catch (error) {
    console.error("Invalid JWT token: Failed to parse payload", error);
    return undefined;
  }

  return token;
}

/**
 * Graph service for HTTP mode: uses a per-session token accessor instead of MSAL cache.
 */
export class SessionGraphService implements IGraphService {
  private client: Client;

  constructor(private tokenAccessor: () => Promise<string>) {
    this.client = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: () => this.tokenAccessor(),
      },
    });
  }

  async getClient(): Promise<Client> {
    return this.client;
  }

  async getAuthStatus(): Promise<AuthStatus> {
    try {
      const me = await this.client.api("/me").get();
      return {
        isAuthenticated: true,
        userPrincipalName: me?.userPrincipalName ?? undefined,
        displayName: me?.displayName ?? undefined,
      };
    } catch {
      return { isAuthenticated: false };
    }
  }

  isAuthenticated(): boolean {
    return true;
  }

  async startDeviceCodeAuth(): Promise<never> {
    throw new Error(
      "Device code authentication is not supported in HTTP mode. Authentication is handled via OAuth."
    );
  }

  validateToken(token: string): string | undefined {
    return validateGraphToken(token);
  }
}

/**
 * Graph service for stdio mode: singleton with MSAL-backed token cache.
 */
export class GraphService implements IGraphService {
  private static instance: GraphService;
  private client: Client | undefined;
  private isInitialized = false;
  private tokenExpiresAt: Date | undefined;
  private msalApp: PublicClientApplication | undefined;
  private msalAccount: AccountInfo | undefined;

  static getInstance(): GraphService {
    if (!GraphService.instance) {
      GraphService.instance = new GraphService();
    }
    return GraphService.instance;
  }

  private async initializeClient(): Promise<void> {
    if (this.isInitialized) return;

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
    } catch (error) {
      console.error("Failed to initialize Graph client:", error);
    }
  }

  private async acquireToken(): Promise<string> {
    if (!this.msalApp || !this.msalAccount) {
      throw new Error("MSAL not initialized");
    }

    const result = await this.msalApp.acquireTokenSilent({
      scopes: DELEGATED_SCOPES,
      account: this.msalAccount,
    });

    if (!result) {
      throw new Error(
        "Failed to acquire access token. Run /ms-teams:authenticate in Claude Code to sign in again."
      );
    }

    this.tokenExpiresAt = result.expiresOn ?? undefined;
    return result.accessToken;
  }

  async getAuthStatus(): Promise<AuthStatus> {
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
    } catch (error) {
      console.error("Error getting user info:", error);
      return { isAuthenticated: false };
    }
  }

  async getClient(): Promise<Client> {
    await this.initializeClient();

    if (!this.client) {
      throw new Error("Not authenticated. Run /ms-teams:authenticate in Claude Code to sign in.");
    }
    return this.client;
  }

  isAuthenticated(): boolean {
    return !!this.client && this.isInitialized;
  }

  async startDeviceCodeAuth(): Promise<{
    verificationUri: string;
    userCode: string;
    expiresIn: number;
  }> {
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

      const pendingAuth = msalApp.acquireTokenByDeviceCode({
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
      pendingAuth.catch(reject);

      // When auth completes successfully, re-initialize the Graph client
      pendingAuth
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

  validateToken(token: string): string | undefined {
    return validateGraphToken(token);
  }
}
