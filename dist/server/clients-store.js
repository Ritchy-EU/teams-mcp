import { mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { DATA_DIR } from "../config.js";
const CLIENTS_PATH = join(DATA_DIR, "oauth-clients.json");
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
export class FileClientsStore {
    path;
    clients = new Map();
    constructor(path = CLIENTS_PATH) {
        this.path = path;
        this.load();
    }
    getClient(clientId) {
        return this.clients.get(clientId);
    }
    registerClient(client) {
        const full = {
            ...client,
            client_id: crypto.randomUUID(),
            client_id_issued_at: Math.floor(Date.now() / 1000),
        };
        this.clients.set(full.client_id, full);
        this.persist();
        return full;
    }
    get size() {
        return this.clients.size;
    }
    load() {
        let raw;
        try {
            raw = readFileSync(this.path, "utf8");
        }
        catch (error) {
            if (error.code !== "ENOENT") {
                console.error(`[OAuth] Could not read client registry at ${this.path}:`, error);
            }
            return;
        }
        try {
            const parsed = JSON.parse(raw);
            for (const client of parsed) {
                if (client?.client_id) {
                    this.clients.set(client.client_id, client);
                }
            }
            console.log(`[OAuth] Loaded ${this.clients.size} registered client(s) from ${this.path}`);
        }
        catch (error) {
            // A corrupt registry must not take the server down: clients can re-register.
            console.error(`[OAuth] Client registry at ${this.path} is corrupt, ignoring it:`, error);
        }
    }
    persist() {
        const data = JSON.stringify([...this.clients.values()], null, 2);
        const tmp = `${this.path}.tmp`;
        try {
            mkdirSync(dirname(this.path), { recursive: true });
            // Write-then-rename so a crash mid-write cannot truncate the registry.
            writeFileSync(tmp, data, { encoding: "utf8", mode: 0o600 });
            renameSync(tmp, this.path);
        }
        catch (error) {
            console.error(`[OAuth] Could not persist client registry to ${this.path}:`, error);
        }
    }
}
export { CLIENTS_PATH };
//# sourceMappingURL=clients-store.js.map