import http from "node:http";
import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { PORT } from "../config.js";
import { GraphService } from "../services/graph.js";
import { registerAuthTools } from "../tools/auth.js";
import { registerChatTools } from "../tools/chats.js";
import { registerOrganizationTools } from "../tools/organization.js";
import { registerSearchTools } from "../tools/search.js";
import { registerTeamsTools } from "../tools/teams.js";
import { registerUsersTools } from "../tools/users.js";

const sessions = new Map<string, StreamableHTTPServerTransport>();

function createMcpServer(): McpServer {
  const server = new McpServer({
    name: "teams-mcp",
    version: "0.7.0",
  });

  const graphService = GraphService.getInstance();

  registerAuthTools(server, graphService);
  registerUsersTools(server, graphService);
  registerOrganizationTools(server, graphService);
  registerTeamsTools(server, graphService);
  registerChatTools(server, graphService);
  registerSearchTools(server, graphService);

  return server;
}

async function readBody(req: http.IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    let data = "";
    req.on("data", (chunk) => { data += chunk; });
    req.on("end", () => {
      try {
        resolve(data ? JSON.parse(data) : undefined);
      } catch {
        resolve(undefined);
      }
    });
    req.on("error", reject);
  });
}

export async function startHttpServer(): Promise<void> {
  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url ?? "/", `http://localhost`);

    if (url.pathname === "/health" && req.method === "GET") {
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ status: "ok", service: "teams-mcp", version: "0.7.0" }));
      return;
    }

    if (url.pathname === "/mcp") {
      const sessionId = req.headers["mcp-session-id"] as string | undefined;

      if (req.method === "DELETE") {
        if (sessionId && sessions.has(sessionId)) {
          const transport = sessions.get(sessionId)!;
          await transport.close();
          sessions.delete(sessionId);
          res.writeHead(200, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ message: "Session closed" }));
        } else {
          res.writeHead(404, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ error: "Session not found" }));
        }
        return;
      }

      if (req.method === "POST" || req.method === "GET") {
        const body = req.method === "POST" ? await readBody(req) : undefined;

        let transport: StreamableHTTPServerTransport;

        if (sessionId && sessions.has(sessionId)) {
          transport = sessions.get(sessionId)!;
        } else if (!sessionId && isInitializeRequest(body)) {
          const newSessionId = randomUUID();

          transport = new StreamableHTTPServerTransport({
            sessionIdGenerator: () => newSessionId,
            onsessioninitialized: (sid) => {
              sessions.set(sid, transport);
            },
          });

          transport.onclose = () => {
            sessions.delete(newSessionId);
          };

          // Cast needed due to exactOptionalPropertyTypes mismatch in SDK types
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const mcpServer = createMcpServer();
          await mcpServer.connect(transport as any);
        } else {
          res.writeHead(400, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ error: "Bad Request: missing or invalid session" }));
          return;
        }

        await transport.handleRequest(req, res, body);
        return;
      }

      res.writeHead(405, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: "Method not allowed" }));
      return;
    }

    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "Not found" }));
  });

  server.listen(PORT, "0.0.0.0", () => {
    console.error(`Teams MCP HTTP server listening on port ${PORT}`);
    console.error(`MCP endpoint: http://0.0.0.0:${PORT}/mcp`);
    console.error(`Health: http://0.0.0.0:${PORT}/health`);
  });
}
