# Copilot Instructions — teams-mcp

## Architecture

This is an MCP (Model Context Protocol) server that bridges AI assistants to Microsoft Graph APIs for Teams, chats, users, and search. It runs as a CLI tool (`npx @ritchy-eu/teams-mcp@latest`) over **stdio transport**.

### Key components

- **`src/index.ts`** — Entry point. Dual-mode: CLI commands (`authenticate`, `check`, `logout`) and MCP server mode (default). Registers all tool modules against a single `McpServer` instance.
- **`src/services/graph.ts`** — `GraphService` singleton. Manages Microsoft Graph client initialization from a stored OAuth token (`~/.msgraph-mcp-auth.json`). All Graph API calls flow through `graphService.getClient()`.
- **`src/tools/*.ts`** — Tool registration modules (`auth`, `chats`, `search`, `teams`, `users`). Each exports a `register*Tools(server, graphService)` function that calls `server.tool()` with Zod schemas for input validation.
- **`src/types/graph.ts`** — Re-exports `@microsoft/microsoft-graph-types` and defines custom response/summary interfaces (`GraphApiResponse<T>`, `*Summary` types). All API response shapes live here.
- **`src/utils/`** — Shared helpers: `markdown.ts` (markdown→sanitized HTML via marked+DOMPurify), `attachments.ts` (image hosted content uploads), `users.ts` (@mention parsing and user lookup).

### Data flow

```
AI Assistant → stdio → McpServer → register*Tools handler → GraphService.getClient() → Microsoft Graph API
```

## Development

```bash
npm run build          # Clean + tsc compile
npm run dev            # Watch mode (node --watch)
npm test               # vitest run
npm run test:watch     # vitest interactive
npm run test:coverage  # vitest with v8 coverage (80% threshold on all metrics)
npm run lint           # biome check
npm run lint:fix       # biome check --write --unsafe
```

## TypeScript & Module Conventions

- **ESM-only** (`"type": "module"` in package.json). Use `.js` extensions in import paths (TypeScript `moduleResolution: "bundler"`).
- **Strict mode** enabled with `noUnusedLocals`, `noUnusedParameters`, `exactOptionalPropertyTypes`, `noImplicitOverride`.
- Prefix unused catch variables with underscore: `catch (_error)`.
- Linting/formatting via **Biome** (not ESLint/Prettier). Use `node:` protocol for Node.js built-in imports.

## Testing Patterns

- **Framework**: Vitest with global test APIs (`describe`, `it`, `expect` — no imports needed).
- **HTTP mocking**: MSW (Mock Service Worker) intercepts `graph.microsoft.com` calls. Shared mock handlers and fixtures live in `src/test-utils/setup.ts`. Global setup in `src/test-utils/vitest.setup.ts` starts/resets/stops the MSW server.
- **Tool tests**: Mock `GraphService` and `McpServer` with `vi.fn()`, then extract registered handler functions from `mockServer.tool.mock.calls` to invoke directly. See `src/tools/__tests__/chats.test.ts` for the canonical pattern:
  ```typescript
  const call = vi.mocked(mockServer.tool).mock.calls.find(([name]) => name === "tool_name");
  const handler = call?.[3] as (args: any) => Promise<any>;
  ```
- **Coverage**: 80% threshold (branches, functions, lines, statements). `index.ts` and `test-utils/` are excluded from coverage.

## Tool Registration Pattern

Every tool module follows this structure — maintain it when adding new tools:

```typescript
export function registerXxxTools(server: McpServer, graphService: GraphService) {
  server.tool(
    "tool_name",           // snake_case tool name
    "Description string",  // user-facing description
    { /* Zod schema */ },  // input validation
    async (args) => {
      const client = await graphService.getClient();
      // ... Graph API call ...
      return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
    }
  );
}
```

After creating a new tool module, register it in `src/index.ts` inside `startMcpServer()`.

## API Response Conventions

- Return JSON-stringified results as `{ content: [{ type: "text", text: ... }] }`.
- Error responses: catch errors, extract message via `error instanceof Error ? error.message : "Unknown error occurred"`, return `❌ Error: ${errorMessage}`.
- All Graph response types use optional properties with `| undefined` to handle API variability (see `*Summary` interfaces in `src/types/graph.ts`).
- Message sending tools support a `format` parameter (`"text"` | `"markdown"`). When `"markdown"`, content is converted via `markdownToHtml()` and sent as `contentType: "html"`.
