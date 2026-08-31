import type { OAuthRegisteredClientsStore } from "@modelcontextprotocol/sdk/server/auth/clients.js";
import type { OAuthClientInformationFull } from "@modelcontextprotocol/sdk/shared/auth.js";
declare const CLIENTS_PATH: string;
/**
 * File-backed store for dynamically registered MCP clients.
 *
 * MCP clients register once via DCR and keep their `client_id` locally, alongside the
 * Microsoft refresh token. If the server forgets that registration, every subsequent
 * `POST /oauth/token` is rejected with `invalid_client` — the refresh token is still
 * valid, but unusable — and the user has to sign in again. Keeping the registry on disk
 * is what makes a restart survivable.
 *
 * Registrations are held in memory and written through to disk, so reads never touch
 * the filesystem and cannot race with a concurrent write.
 */
export declare class FileClientsStore implements OAuthRegisteredClientsStore {
    private readonly path;
    private clients;
    constructor(path?: string);
    getClient(clientId: string): OAuthClientInformationFull | undefined;
    registerClient(client: Omit<OAuthClientInformationFull, "client_id" | "client_id_issued_at">): OAuthClientInformationFull;
    get size(): number;
    private load;
    private persist;
}
export { CLIENTS_PATH };
//# sourceMappingURL=clients-store.d.ts.map