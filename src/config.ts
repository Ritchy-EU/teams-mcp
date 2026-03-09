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
export const DELEGATED_SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "User.Read.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "TeamMember.Read.All",
  "Chat.ReadBasic",
  "Chat.ReadWrite",
  "Files.Read.All",
];

// HTTP server configuration (used in `serve` mode)
export const PORT = parseInt(process.env.PORT || "3000", 10);
export const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;
export const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
