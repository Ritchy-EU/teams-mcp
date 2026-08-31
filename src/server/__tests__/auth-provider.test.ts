import {
  InvalidGrantError,
  InvalidTokenError,
  ServerError,
} from "@modelcontextprotocol/sdk/server/auth/errors.js";
import type { AuthorizationParams } from "@modelcontextprotocol/sdk/server/auth/provider.js";
import type { OAuthClientInformationFull } from "@modelcontextprotocol/sdk/shared/auth.js";
import type { Response } from "express";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MicrosoftEntraOAuthProvider } from "../auth-provider.js";

/** Builds an unsigned JWT with the given payload — enough for validateGraphToken. */
function makeToken(payload: Record<string, unknown>): string {
  const encode = (obj: Record<string, unknown>) => btoa(JSON.stringify(obj));
  return `${encode({ alg: "none", typ: "JWT" })}.${encode(payload)}.signature`;
}

const GRAPH_AUDIENCE = "https://graph.microsoft.com";

function validPayload(overrides: Record<string, unknown> = {}) {
  return {
    aud: GRAPH_AUDIENCE,
    iss: "https://login.microsoftonline.com/test-tenant-id/v2.0",
    exp: Math.floor(Date.now() / 1000) + 3600,
    appid: "test-app-id",
    scp: "User.Read Chat.Read",
    ...overrides,
  };
}

const stubClient = { client_id: "mcp-client" } as OAuthClientInformationFull;

const stubClientsStore = {
  getClient: () => undefined,
  registerClient: (client: Omit<OAuthClientInformationFull, "client_id" | "client_id_issued_at">) =>
    ({ ...client, client_id: "mcp-client", client_id_issued_at: 0 }) as OAuthClientInformationFull,
};

describe("MicrosoftEntraOAuthProvider", () => {
  let provider: MicrosoftEntraOAuthProvider;
  let logSpy: ReturnType<typeof vi.spyOn>;
  let errorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    provider = new MicrosoftEntraOAuthProvider(stubClientsStore);
    logSpy = vi.spyOn(console, "log").mockImplementation(() => {
      // Keep the provider's diagnostic logging out of the test output.
    });
    errorSpy = vi.spyOn(console, "error").mockImplementation(() => {
      // Expected on the failure paths asserted below.
    });
  });

  afterEach(() => {
    logSpy.mockRestore();
    errorSpy.mockRestore();
  });

  describe("verifyAccessToken", () => {
    it("accepts a valid Graph token and reports its metadata", async () => {
      const exp = Math.floor(Date.now() / 1000) + 3600;
      const info = await provider.verifyAccessToken(makeToken(validPayload({ exp })));

      expect(info.clientId).toBe("test-app-id");
      expect(info.expiresAt).toBe(exp);
      expect(info.scopes).toEqual(["User.Read", "Chat.Read"]);
    });

    it("throws InvalidTokenError for an expired token so the client gets a 401", async () => {
      // requireBearerAuth only maps InvalidTokenError to 401 + WWW-Authenticate. Any other
      // error becomes a 500, and clients retry with the same expired token forever instead
      // of refreshing it.
      const expired = makeToken(validPayload({ exp: Math.floor(Date.now() / 1000) - 60 }));

      await expect(provider.verifyAccessToken(expired)).rejects.toThrow(InvalidTokenError);
    });

    it("throws InvalidTokenError for a malformed token", async () => {
      await expect(provider.verifyAccessToken("not-a-jwt")).rejects.toThrow(InvalidTokenError);
    });

    it("throws InvalidTokenError when the audience is not Microsoft Graph", async () => {
      const wrongAudience = makeToken(validPayload({ aud: "https://example.com" }));

      await expect(provider.verifyAccessToken(wrongAudience)).rejects.toThrow(InvalidTokenError);
    });
  });

  describe("exchangeRefreshToken", () => {
    const stubTokens = {
      access_token: "at",
      token_type: "Bearer",
      expires_in: 3600,
      refresh_token: "new-rt",
      scope: "User.Read",
    };

    function mockTokenEndpoint(status: number, body: unknown) {
      return vi.spyOn(globalThis, "fetch").mockResolvedValue(
        new Response(typeof body === "string" ? body : JSON.stringify(body), {
          status,
          headers: { "Content-Type": "application/json" },
        })
      );
    }

    afterEach(() => {
      vi.restoreAllMocks();
    });

    it("always sends offline_access so Entra keeps rotating the refresh token", async () => {
      const fetchSpy = mockTokenEndpoint(200, stubTokens);

      await provider.exchangeRefreshToken(stubClient, "rt", ["User.Read"]);

      const body = String(fetchSpy.mock.calls[0]?.[1]?.body);
      expect(new URLSearchParams(body).get("scope")?.split(" ")).toContain("offline_access");
    });

    it("surfaces a dead refresh token as invalid_grant, not a server error", async () => {
      // The SDK turns a plain Error into 500 server_error, which makes clients retry a
      // token that will never work again instead of starting a fresh authorization.
      mockTokenEndpoint(400, {
        error: "invalid_grant",
        error_description: "AADSTS700082: The refresh token has expired.",
      });

      await expect(provider.exchangeRefreshToken(stubClient, "dead-rt")).rejects.toThrow(
        InvalidGrantError
      );
    });

    it("maps an upstream outage to a server error", async () => {
      mockTokenEndpoint(503, { error: "temporarily_unavailable" });

      await expect(provider.exchangeRefreshToken(stubClient, "rt")).rejects.toThrow(ServerError);
    });

    it("maps a non-JSON upstream failure without throwing", async () => {
      mockTokenEndpoint(502, "<html>bad gateway</html>");

      await expect(provider.exchangeRefreshToken(stubClient, "rt")).rejects.toThrow(ServerError);
    });

    it("returns the rotated refresh token on success", async () => {
      mockTokenEndpoint(200, stubTokens);

      const tokens = await provider.exchangeRefreshToken(stubClient, "rt");

      expect(tokens.refresh_token).toBe("new-rt");
      expect(tokens.access_token).toBe("at");
    });
  });

  describe("authorize", () => {
    function capturingResponse() {
      const redirects: string[] = [];
      const res = { redirect: (url: string) => redirects.push(url) } as unknown as Response;
      return { res, redirects };
    }

    const baseParams = {
      redirectUri: "http://localhost:40056/callback",
      codeChallenge: "challenge",
      state: "state-value",
    };

    it("always requests offline_access, even when the client omits scopes", async () => {
      const { res, redirects } = capturingResponse();

      await provider.authorize(stubClient, baseParams as AuthorizationParams, res);

      const scope = new URL(redirects[0]).searchParams.get("scope")?.split(" ");
      expect(scope).toContain("offline_access");
      // Without a refresh token the session dies unrecoverably after ~1h.
      expect(scope).toContain("User.Read");
    });

    it("adds offline_access to a narrow client-requested scope set", async () => {
      const { res, redirects } = capturingResponse();

      await provider.authorize(
        stubClient,
        { ...baseParams, scopes: ["User.Read"] } as AuthorizationParams,
        res
      );

      expect(new URL(redirects[0]).searchParams.get("scope")).toContain("offline_access");
    });

    it("sends our own callback URL to Microsoft, not the client's", async () => {
      const { res, redirects } = capturingResponse();

      await provider.authorize(stubClient, baseParams as AuthorizationParams, res);

      const params = new URL(redirects[0]).searchParams;
      expect(params.get("redirect_uri")).not.toBe(baseParams.redirectUri);
      expect(params.get("redirect_uri")).toContain("/oauth/callback");
      expect(params.get("state")).toBe("state-value");
    });

    it("remembers the client's redirect_uri for the callback", async () => {
      const { res } = capturingResponse();

      await provider.authorize(stubClient, baseParams as AuthorizationParams, res);

      expect(provider.handleCallback("state-value")).toBe(baseParams.redirectUri);
      expect(provider.handleCallback("unknown-state")).toBeUndefined();
    });
  });
});
