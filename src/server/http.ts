import { randomUUID } from "node:crypto";
import type { ErrorRequestHandler } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import type { Transport } from "@modelcontextprotocol/sdk/shared/transport.js";
import { mcpAuthRouter } from "@modelcontextprotocol/sdk/server/auth/router.js";
import { requireBearerAuth } from "@modelcontextprotocol/sdk/server/auth/middleware/bearerAuth.js";
import { createMcpExpressApp } from "@modelcontextprotocol/sdk/server/express.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { SessionGraphService } from "../services/graph.js";
import { registerAuthTools } from "../tools/auth.js";
import { registerChatTools } from "../tools/chats.js";
import { registerSearchTools } from "../tools/search.js";
import { registerTeamsTools } from "../tools/teams.js";
import { registerOrganizationTools } from "../tools/organization.js";
import { registerUsersTools } from "../tools/users.js";
import { MicrosoftEntraOAuthProvider } from "./auth-provider.js";
import { SessionManager } from "./session-manager.js";
import { PORT, BASE_URL, DELEGATED_SCOPES } from "../config.js";

const HTTP_SCOPES = [
  "offline_access", // Enables refresh tokens for long-lived sessions
  ...DELEGATED_SCOPES,
];

function createSessionServer(tokenAccessor: () => Promise<string>): {
  server: McpServer;
  graphService: SessionGraphService;
} {
  const server = new McpServer({
    name: "teams-mcp",
    version: "0.7.0",
  });

  const graphService = new SessionGraphService(tokenAccessor);

  registerAuthTools(server, graphService);
  registerUsersTools(server, graphService);
  registerOrganizationTools(server, graphService);
  registerTeamsTools(server, graphService);
  registerChatTools(server, graphService);
  registerSearchTools(server, graphService);

  return { server, graphService };
}

export async function startHttpServer(): Promise<void> {
  const sessionManager = new SessionManager();
  const provider = new MicrosoftEntraOAuthProvider();

  const app = createMcpExpressApp({ host: "0.0.0.0" });

  // Trust reverse proxy (nginx, Cloudflare, etc.) for correct client IP detection.
  // Required by express-rate-limit when X-Forwarded-For headers are present.
  app.set("trust proxy", 1);

  // OAuth callback: Microsoft redirects here after user login.
  // We then forward the authorization code to the MCP client's original redirect_uri.
  // This MUST be registered before mcpAuthRouter to avoid route conflicts.
  app.get("/oauth/callback", (req, res) => {
    console.log(`[OAuth] Callback hit: ${req.originalUrl}`);
    const { code, state, error, error_description } = req.query;

    if (error) {
      console.log(`[OAuth] Callback error: ${error} - ${error_description}`);
      res.status(400).send(`Authentication error: ${error_description || error}`);
      return;
    }

    if (!state || typeof state !== "string" || !code || typeof code !== "string") {
      console.log(`[OAuth] Callback missing params: state=${typeof state}, code=${typeof code}`);
      res.status(400).send("Missing authorization code or state parameter");
      return;
    }

    // Look up the MCP client's original redirect_uri
    const originalRedirectUri = provider.handleCallback(state);
    if (!originalRedirectUri) {
      res.status(400).send("Unknown or expired authorization flow. Please try again.");
      return;
    }

    // Redirect to MCP client with the authorization code
    const redirectUrl = new URL(originalRedirectUri);
    redirectUrl.searchParams.set("code", code);
    redirectUrl.searchParams.set("state", state);
    console.log(`[OAuth] Redirecting to MCP client: ${redirectUrl.toString().substring(0, 80)}...`);
    res.redirect(redirectUrl.toString());
  });

  // OAuth endpoints: .well-known metadata, authorize, token, register
  app.use(
    mcpAuthRouter({
      provider,
      issuerUrl: new URL(BASE_URL),
      scopesSupported: HTTP_SCOPES,
      resourceName: "Teams MCP Server",
    })
  );

  const authMiddleware = requireBearerAuth({
    verifier: provider,
    requiredScopes: [],
  });

  // Health check
  app.get("/health", (_req, res) => {
    res.json({
      status: "ok",
      activeSessions: sessionManager.size,
    });
  });

  // POST /mcp — main MCP request handler
  app.post("/mcp", authMiddleware, async (req, res) => {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    const method = req.body?.method ?? "unknown";
    console.log(`[MCP] POST /mcp method=${method} sessionId=${sessionId ?? "none"}`);

    try {
      // Existing session
      if (sessionId && sessionManager.has(sessionId)) {
        const session = sessionManager.get(sessionId)!;
        // Update the session's token from the current request — the MCP client
        // may have refreshed it since the session was created.
        session.updateToken(req.auth!.token);
        await session.transport.handleRequest(req, res, req.body);
        return;
      }

      // New session initialization
      if (!sessionId && isInitializeRequest(req.body)) {
        let currentToken = req.auth!.token;
        console.log(`[MCP] New session initialization, token length=${currentToken.length}`);

        const { server: mcpServer, graphService } = createSessionServer(
          async () => currentToken
        );

        const transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID(),
          onsessioninitialized: (sid) => {
            sessionManager.set(sid, {
              transport,
              server: mcpServer,
              graphService,
              lastActivity: Date.now(),
              updateToken: (newToken: string) => {
                currentToken = newToken;
              },
            });
          },
        });

        transport.onclose = () => {
          const sid = transport.sessionId;
          if (sid) {
            sessionManager.delete(sid);
          }
        };

        await mcpServer.connect(transport as unknown as Transport);
        await transport.handleRequest(req, res, req.body);
        return;
      }

      // Invalid request
      res.status(400).json({
        jsonrpc: "2.0",
        error: {
          code: -32000,
          message: "Bad Request: No valid session ID provided",
        },
        id: null,
      });
    } catch (error) {
      console.error("Error handling MCP request:", error);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        });
      }
    }
  });

  // GET /mcp — SSE streams
  app.get("/mcp", authMiddleware, async (req, res) => {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    if (!sessionId || !sessionManager.has(sessionId)) {
      res.status(400).send("Invalid or missing session ID");
      return;
    }
    const session = sessionManager.get(sessionId)!;
    await session.transport.handleRequest(req, res);
  });

  // DELETE /mcp — session termination
  app.delete("/mcp", authMiddleware, async (req, res) => {
    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    if (!sessionId || !sessionManager.has(sessionId)) {
      res.status(400).send("Invalid or missing session ID");
      return;
    }
    try {
      const session = sessionManager.get(sessionId)!;
      await session.transport.handleRequest(req, res);
    } catch (error) {
      console.error("Error handling session termination:", error);
      if (!res.headersSent) {
        res.status(500).send("Error processing session termination");
      }
    }
  });

  // Global JSON error handler — prevents Express 5 from returning HTML error pages.
  // Must be registered AFTER all routes.
  const jsonErrorHandler: ErrorRequestHandler = (err, _req, res, _next) => {
    console.error("[Express] Unhandled error:", err);
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: "2.0",
        error: {
          code: -32603,
          message: "Internal server error",
        },
        id: null,
      });
    }
  };
  app.use(jsonErrorHandler);

  app.listen(PORT, () => {
    console.log(`Teams MCP HTTP Server listening on port ${PORT}`);
    console.log(`Base URL: ${BASE_URL}`);
    console.log(`Health check: ${BASE_URL}/health`);
  });

  // Graceful shutdown
  const shutdown = async () => {
    console.log("Shutting down server...");
    await sessionManager.shutdown();
    console.log("Server shutdown complete");
    process.exit(0);
  };

  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);
}
