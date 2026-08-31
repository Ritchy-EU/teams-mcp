import { mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import type { OAuthRegisteredClientsStore } from "@modelcontextprotocol/sdk/server/auth/clients.js";
import type { OAuthClientInformationFull } from "@modelcontextprotocol/sdk/shared/auth.js";
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
export class FileClientsStore implements OAuthRegisteredClientsStore {
  private clients = new Map<string, OAuthClientInformationFull>();

  constructor(private readonly path: string = CLIENTS_PATH) {
    this.load();
  }

  getClient(clientId: string): OAuthClientInformationFull | undefined {
    return this.clients.get(clientId);
  }

  registerClient(
    client: Omit<OAuthClientInformationFull, "client_id" | "client_id_issued_at">
  ): OAuthClientInformationFull {
    const full: OAuthClientInformationFull = {
      ...client,
      client_id: crypto.randomUUID(),
      client_id_issued_at: Math.floor(Date.now() / 1000),
    } as OAuthClientInformationFull;
    this.clients.set(full.client_id, full);
    this.persist();
    return full;
  }

  get size(): number {
    return this.clients.size;
  }

  private load(): void {
    let raw: string;
    try {
      raw = readFileSync(this.path, "utf8");
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") {
        console.error(`[OAuth] Could not read client registry at ${this.path}:`, error);
      }
      return;
    }

    try {
      const parsed = JSON.parse(raw) as OAuthClientInformationFull[];
      for (const client of parsed) {
        if (client?.client_id) {
          this.clients.set(client.client_id, client);
        }
      }
      console.log(`[OAuth] Loaded ${this.clients.size} registered client(s) from ${this.path}`);
    } catch (error) {
      // A corrupt registry must not take the server down: clients can re-register.
      console.error(`[OAuth] Client registry at ${this.path} is corrupt, ignoring it:`, error);
    }
  }

  private persist(): void {
    const data = JSON.stringify([...this.clients.values()], null, 2);
    const tmp = `${this.path}.tmp`;
    try {
      mkdirSync(dirname(this.path), { recursive: true });
      // Write-then-rename so a crash mid-write cannot truncate the registry.
      writeFileSync(tmp, data, { encoding: "utf8", mode: 0o600 });
      renameSync(tmp, this.path);
    } catch (error) {
      console.error(`[OAuth] Could not persist client registry to ${this.path}:`, error);
    }
  }
}

export { CLIENTS_PATH };
