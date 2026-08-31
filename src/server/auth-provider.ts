import type { OAuthRegisteredClientsStore } from "@modelcontextprotocol/sdk/server/auth/clients.js";
import {
  InvalidGrantError,
  InvalidRequestError,
  InvalidTokenError,
  ServerError,
} from "@modelcontextprotocol/sdk/server/auth/errors.js";
import type {
  AuthorizationParams,
  OAuthServerProvider,
} from "@modelcontextprotocol/sdk/server/auth/provider.js";
import type { AuthInfo } from "@modelcontextprotocol/sdk/server/auth/types.js";
import type {
  OAuthClientInformationFull,
  OAuthTokenRevocationRequest,
  OAuthTokens,
} from "@modelcontextprotocol/sdk/shared/auth.js";
import type { Response } from "express";
import { AZURE_CLIENT_SECRET, BASE_URL, CLIENT_ID, HTTP_SCOPES, TENANT_ID } from "../config.js";
import { validateGraphToken } from "../services/graph.js";
import { FileClientsStore } from "./clients-store.js";

const TOKEN_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const AUTHORIZATION_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`;

/** Our server's callback URL registered in Azure AD */
const OUR_CALLBACK_URL = `${BASE_URL}/oauth/callback`;

/**
 * Translates a failed Entra token response into an OAuth error.
 *
 * The SDK's token handler only maps `OAuthError` subclasses to a 400 with the real error
 * code; anything else becomes a 500 `server_error`. That distinction matters: a client
 * told `invalid_grant` discards its refresh token and starts a fresh authorization, while
 * a client told `server_error` just retries — forever, if the token is genuinely dead.
 */
function upstreamTokenError(status: number, body: string): Error {
  let code: string | undefined;
  let description: string | undefined;
  try {
    const parsed = JSON.parse(body);
    code = typeof parsed.error === "string" ? parsed.error : undefined;
    description =
      typeof parsed.error_description === "string" ? parsed.error_description : undefined;
  } catch {
    // Not JSON — fall through to the status-based mapping below.
  }

  const message = description ?? body ?? `Token request failed (${status})`;

  if (code === "invalid_grant") {
    return new InvalidGrantError(message);
  }
  if (status >= 500) {
    return new ServerError(message);
  }
  if (status >= 400) {
    return new InvalidRequestError(message);
  }
  return new ServerError(message);
}

/** TTL for pending auth flows (10 minutes) */
const PENDING_AUTH_TTL_MS = 10 * 60 * 1000;

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

  private _clientsStore: OAuthRegisteredClientsStore;

  /**
   * Maps OAuth state → MCP client's original redirect_uri.
   * Entries are cleaned up after use or after TTL expiry.
   */
  private pendingAuthFlows = new Map<string, PendingAuthFlow>();

  constructor(clientsStore: OAuthRegisteredClientsStore = new FileClientsStore()) {
    this._clientsStore = clientsStore;
  }

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

    // Always request our full scope set plus `offline_access`. MCP clients may send a
    // narrower `scope` (or none at all, in which case params.scopes is undefined), and
    // any request without `offline_access` gets no refresh token from Entra — leaving
    // the session unrecoverable once the access token expires.
    const scopes = new Set(params.scopes?.length ? params.scopes : []);
    for (const scope of HTTP_SCOPES) scopes.add(scope);
    scopes.add("offline_access");
    searchParams.set("scope", [...scopes].join(" "));

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
          `knownStates=[${[...this.pendingAuthFlows.keys()].map((s) => `${s.substring(0, 8)}...`).join(", ")}]`
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
      throw upstreamTokenError(response.status, errorBody);
    }

    const data = await response.json();
    console.log(
      `[OAuth] Token exchange successful: hasAccessToken=${!!data.access_token}, ` +
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
    // Entra rotates refresh tokens on every redemption, but only returns a new one when
    // `offline_access` is in the request. Dropping it here would break the chain after a
    // single refresh. Widen only by that scope — re-requesting scopes the user never
    // consented to would fail the grant outright.
    const scopeSet = new Set(scopes?.length ? scopes : HTTP_SCOPES);
    scopeSet.add("offline_access");
    params.set("scope", [...scopeSet].join(" "));

    const response = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    });

    if (!response.ok) {
      const errorBody = await response.text();
      console.error(`[OAuth] Token refresh failed (${response.status}): ${errorBody}`);
      // A dead refresh token must reach the client as `invalid_grant` so it re-authorizes
      // instead of retrying a token that will never work again.
      throw upstreamTokenError(response.status, errorBody);
    }

    const data = await response.json();
    if (!data.refresh_token) {
      console.error(
        "[OAuth] Token refresh returned no new refresh token — the client will not be able " +
          "to refresh again. Check that `offline_access` is consented for this app registration."
      );
    }
    console.log(
      `[OAuth] Token refreshed: expiresIn=${data.expires_in}s, ` +
        `rotatedRefreshToken=${!!data.refresh_token}, scope=${data.scope}`
    );
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
      // Must be an InvalidTokenError: requireBearerAuth maps it to 401 with a
      // WWW-Authenticate header, which is what makes the MCP client refresh. Any other
      // error becomes a 500, and clients just retry with the same expired token.
      throw new InvalidTokenError("Invalid or expired Microsoft Graph token");
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
