import type { Response } from "express";
import type {
  OAuthClientInformationFull,
  OAuthTokenRevocationRequest,
  OAuthTokens,
} from "@modelcontextprotocol/sdk/shared/auth.js";
import type {
  AuthorizationParams,
  OAuthServerProvider,
} from "@modelcontextprotocol/sdk/server/auth/provider.js";
import type { OAuthRegisteredClientsStore } from "@modelcontextprotocol/sdk/server/auth/clients.js";
import type { AuthInfo } from "@modelcontextprotocol/sdk/server/auth/types.js";
import { validateGraphToken } from "../services/graph.js";
import { CLIENT_ID, AZURE_CLIENT_SECRET, TENANT_ID, BASE_URL } from "../config.js";

const TOKEN_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const AUTHORIZATION_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`;

/** Our server's callback URL registered in Azure AD */
const OUR_CALLBACK_URL = `${BASE_URL}/oauth/callback`;

/** TTL for pending auth flows (10 minutes) */
const PENDING_AUTH_TTL_MS = 10 * 60 * 1000;

/**
 * In-memory store for dynamically registered MCP clients.
 */
class InMemoryClientsStore implements OAuthRegisteredClientsStore {
  private clients = new Map<string, OAuthClientInformationFull>();

  getClient(clientId: string): OAuthClientInformationFull | undefined {
    return this.clients.get(clientId);
  }

  registerClient(
    client: Omit<OAuthClientInformationFull, "client_id" | "client_id_issued_at">
  ): OAuthClientInformationFull {
    const clientId = crypto.randomUUID();
    const full: OAuthClientInformationFull = {
      ...client,
      client_id: clientId,
      client_id_issued_at: Math.floor(Date.now() / 1000),
    } as OAuthClientInformationFull;
    this.clients.set(clientId, full);
    return full;
  }
}

interface PendingAuthFlow {
  originalRedirectUri: string;
  createdAt: number;
}

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
export class MicrosoftEntraOAuthProvider implements OAuthServerProvider {
  skipLocalPkceValidation = true;

  private _clientsStore = new InMemoryClientsStore();

  /**
   * Maps OAuth state → MCP client's original redirect_uri.
   * Entries are cleaned up after use or after TTL expiry.
   */
  private pendingAuthFlows = new Map<string, PendingAuthFlow>();

  get clientsStore(): OAuthRegisteredClientsStore {
    return this._clientsStore;
  }

  async authorize(
    _client: OAuthClientInformationFull,
    params: AuthorizationParams,
    res: Response
  ): Promise<void> {
    // Store the MCP client's original redirect_uri keyed by state
    if (params.state) {
      this.cleanupExpiredFlows();
      this.pendingAuthFlows.set(params.state, {
        originalRedirectUri: params.redirectUri,
        createdAt: Date.now(),
      });
      console.log(
        `[OAuth] Stored pending auth flow: state=${params.state.substring(0, 8)}..., ` +
        `redirectUri=${params.redirectUri}, ` +
        `pendingFlows=${this.pendingAuthFlows.size}`
      );
    }

    const targetUrl = new URL(AUTHORIZATION_URL);
    const searchParams = new URLSearchParams({
      client_id: CLIENT_ID,
      response_type: "code",
      redirect_uri: OUR_CALLBACK_URL, // OUR registered URL, not the MCP client's
      code_challenge: params.codeChallenge,
      code_challenge_method: "S256",
    });

    if (params.state) searchParams.set("state", params.state);
    if (params.scopes?.length) searchParams.set("scope", params.scopes.join(" "));

    targetUrl.search = searchParams.toString();
    res.redirect(targetUrl.toString());
  }

  /**
   * Called by the /oauth/callback handler to retrieve the MCP client's
   * original redirect_uri for a given state.
   *
   * Note: We intentionally do NOT delete the entry on first access.
   * Reverse proxies (ngrok, Cloudflare Tunnel) may show interstitial pages
   * that cause duplicate requests to /oauth/callback. The entry is cleaned
   * up by the TTL-based cleanup instead.
   */
  handleCallback(state: string): string | undefined {
    const flow = this.pendingAuthFlows.get(state);
    if (!flow) {
      console.log(
        `[OAuth] handleCallback: state=${state.substring(0, 8)}... NOT FOUND. ` +
        `pendingFlows=${this.pendingAuthFlows.size}, ` +
        `knownStates=[${[...this.pendingAuthFlows.keys()].map(s => s.substring(0, 8) + "...").join(", ")}]`
      );
      return undefined;
    }

    // Check TTL
    if (Date.now() - flow.createdAt > PENDING_AUTH_TTL_MS) {
      console.log(
        `[OAuth] handleCallback: state=${state.substring(0, 8)}... EXPIRED ` +
        `(age=${Math.round((Date.now() - flow.createdAt) / 1000)}s)`
      );
      this.pendingAuthFlows.delete(state);
      return undefined;
    }

    console.log(
      `[OAuth] handleCallback: state=${state.substring(0, 8)}... → ${flow.originalRedirectUri}`
    );
    return flow.originalRedirectUri;
  }

  async challengeForAuthorizationCode(
    _client: OAuthClientInformationFull,
    _authorizationCode: string
  ): Promise<string> {
    // Upstream (Microsoft) validates PKCE, not us
    return "";
  }

  async exchangeAuthorizationCode(
    _client: OAuthClientInformationFull,
    authorizationCode: string,
    codeVerifier?: string,
    _redirectUri?: string
  ): Promise<OAuthTokens> {
    console.log(
      `[OAuth] Exchanging authorization code (length=${authorizationCode.length}, ` +
      `hasVerifier=${!!codeVerifier}, redirectUri=${OUR_CALLBACK_URL})`
    );

    const params = new URLSearchParams({
      grant_type: "authorization_code",
      client_id: CLIENT_ID,
      code: authorizationCode,
      // Must match the redirect_uri sent to Microsoft in authorize()
      redirect_uri: OUR_CALLBACK_URL,
    });

    if (AZURE_CLIENT_SECRET) {
      params.append("client_secret", AZURE_CLIENT_SECRET);
    }
    if (codeVerifier) {
      params.append("code_verifier", codeVerifier);
    }

    const response = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    });

    if (!response.ok) {
      const errorBody = await response.text();
      console.error(`[OAuth] Token exchange failed (${response.status}): ${errorBody}`);
      throw new Error(`Token exchange failed (${response.status}): ${errorBody}`);
    }

    const data = await response.json();
    console.log(
      `[OAuth] Token exchange successful: ` +
      `hasAccessToken=${!!data.access_token}, ` +
      `hasRefreshToken=${!!data.refresh_token}, ` +
      `expiresIn=${data.expires_in}s, ` +
      `scope=${data.scope}`
    );
    return {
      access_token: data.access_token,
      token_type: data.token_type ?? "Bearer",
      expires_in: data.expires_in,
      refresh_token: data.refresh_token,
      scope: data.scope,
    };
  }

  async exchangeRefreshToken(
    _client: OAuthClientInformationFull,
    refreshToken: string,
    scopes?: string[]
  ): Promise<OAuthTokens> {
    const params = new URLSearchParams({
      grant_type: "refresh_token",
      client_id: CLIENT_ID,
      refresh_token: refreshToken,
    });

    if (AZURE_CLIENT_SECRET) {
      params.set("client_secret", AZURE_CLIENT_SECRET);
    }
    if (scopes?.length) {
      params.set("scope", scopes.join(" "));
    }

    const response = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    });

    if (!response.ok) {
      const errorBody = await response.text();
      throw new Error(`Token refresh failed (${response.status}): ${errorBody}`);
    }

    const data = await response.json();
    return {
      access_token: data.access_token,
      token_type: data.token_type ?? "Bearer",
      expires_in: data.expires_in,
      refresh_token: data.refresh_token,
      scope: data.scope,
    };
  }

  async verifyAccessToken(token: string): Promise<AuthInfo> {
    console.log(`[OAuth] Verifying access token (length=${token.length})`);
    const validated = validateGraphToken(token);
    if (!validated) {
      console.error("[OAuth] Token verification FAILED");
      throw new Error("Invalid or expired Microsoft Graph token");
    }

    // Decode JWT payload for metadata
    const payload = JSON.parse(atob(token.split(".")[1]));
    console.log(
      `[OAuth] Token verified: clientId=${payload.appid || payload.azp || "unknown"}, ` +
      `exp=${payload.exp ? new Date(payload.exp * 1000).toISOString() : "none"}`
    );

    return {
      token,
      clientId: payload.appid || payload.azp || CLIENT_ID,
      scopes: typeof payload.scp === "string" ? payload.scp.split(" ") : [],
      expiresAt: typeof payload.exp === "number" ? payload.exp : undefined,
    };
  }

  async revokeToken?(
    _client: OAuthClientInformationFull,
    _request: OAuthTokenRevocationRequest
  ): Promise<void> {
    // Microsoft Entra ID doesn't have a standard revocation endpoint
    // Token will expire naturally
  }

  private cleanupExpiredFlows(): void {
    const now = Date.now();
    for (const [state, flow] of this.pendingAuthFlows) {
      if (now - flow.createdAt > PENDING_AUTH_TTL_MS) {
        this.pendingAuthFlows.delete(state);
      }
    }
  }
}
