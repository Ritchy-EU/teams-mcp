import { homedir } from "node:os";
import { join } from "node:path";

// Azure AD application configuration.
//
// IMPORTANT: The built-in CLIENT_ID is the public Microsoft Graph CLI app registration
// (a shared app owned by Microsoft). It is suitable for personal / developer use but
// you should register your own Azure AD app for any production or organisational deployment.
//
// To use your own app registration set the following environment variables:
//   AZURE_CLIENT_ID  — Application (client) ID from Azure Portal > App registrations
//   AZURE_TENANT_ID  — Directory (tenant) ID, or "common" for multi-tenant
//
// To register a new app:
//   1. Azure Portal → Azure Active Directory → App registrations → New registration
//   2. Choose "Public client / native" redirect URI type
//   3. Add the required API permissions (same as DELEGATED_SCOPES in graph.ts)
//   4. Copy the Application (client) ID and Directory (tenant) ID

const clientId = process.env.AZURE_CLIENT_ID;
if (!clientId) {
  throw new Error(
    "AZURE_CLIENT_ID environment variable is required. " +
      "Set it via: claude mcp add --scope user teams-mcp -e AZURE_CLIENT_ID=<your-client-id> -e AZURE_TENANT_ID=<your-tenant-id> -- npx -y github:Ritchy-EU/teams-mcp"
  );
}
export const CLIENT_ID = clientId;

const tenantId = process.env.AZURE_TENANT_ID ?? "common";
export const TENANT_ID = tenantId;
export const AUTHORITY = `https://login.microsoftonline.com/${tenantId}`;

// Scopes for delegated (user) authentication.
// All modes (stdio, HTTP) share this base set of scopes.

/** Scopes sufficient for read-only operations (no sending, no uploads). */
export const READ_ONLY_SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "User.Read.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "TeamMember.Read.All",
  "Chat.Read",
  "Files.Read.All",
];

/** Full scopes including write operations. */
export const FULL_SCOPES = [
  ...READ_ONLY_SCOPES,
  "ChannelMessage.Send",
  "ChannelMessage.ReadWrite",
  "Chat.ReadWrite",
  "ChatMember.ReadWrite",
  "Files.ReadWrite.All",
];

/**
 * Read-only mode: set TEAMS_MCP_READ_ONLY=true (or 1/yes) to skip registering
 * write tools and request only the reduced permission scopes.
 */
export const READ_ONLY = ["1", "true", "yes"].includes(
  (process.env.TEAMS_MCP_READ_ONLY ?? "").toLowerCase()
);

export const DELEGATED_SCOPES = READ_ONLY ? READ_ONLY_SCOPES : FULL_SCOPES;

/**
 * Scopes requested from Microsoft Entra ID in HTTP (OAuth) mode.
 *
 * `offline_access` is what makes Entra issue a refresh token; without it the
 * session dies unrecoverably when the access token expires (~1h), so it is
 * always forced into the outbound request regardless of what the MCP client asked for.
 */
export const HTTP_SCOPES = ["offline_access", ...DELEGATED_SCOPES];

/**
 * Directory for server-side state that must survive a restart (currently the OAuth
 * client registry). In Docker this should be a mounted volume — see docker-compose.yml.
 */
export const DATA_DIR = process.env.TEAMS_MCP_DATA_DIR ?? join(homedir(), ".teams-mcp");

// HTTP server configuration (used in `serve` mode)
export const PORT = Number.parseInt(process.env.PORT || "3000", 10);
export const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;
export const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
