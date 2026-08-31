import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { FileClientsStore } from "../clients-store.js";

describe("FileClientsStore", () => {
  let dir: string;
  let path: string;

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), "teams-mcp-clients-"));
    path = join(dir, "oauth-clients.json");
  });

  afterEach(() => {
    rmSync(dir, { recursive: true, force: true });
  });

  const clientMetadata = {
    client_name: "Claude Code",
    redirect_uris: ["http://localhost:40056/callback"],
  };

  it("returns undefined for an unknown client", () => {
    const store = new FileClientsStore(path);
    expect(store.getClient("nope")).toBeUndefined();
  });

  it("registers a client and returns it by id", () => {
    const store = new FileClientsStore(path);
    const registered = store.registerClient(clientMetadata);

    expect(registered.client_id).toBeTruthy();
    expect(registered.client_id_issued_at).toBeTypeOf("number");
    expect(store.getClient(registered.client_id)).toEqual(registered);
  });

  it("survives a restart: a new store instance still knows the client", () => {
    const registered = new FileClientsStore(path).registerClient(clientMetadata);

    // Simulates the server process being replaced — this is the case that used to
    // invalidate every client's refresh token with `invalid_client`.
    const afterRestart = new FileClientsStore(path);

    expect(afterRestart.getClient(registered.client_id)).toMatchObject({
      client_id: registered.client_id,
      client_name: "Claude Code",
      redirect_uris: ["http://localhost:40056/callback"],
    });
  });

  it("keeps previously registered clients when a new one registers", () => {
    const store = new FileClientsStore(path);
    const first = store.registerClient(clientMetadata);
    const second = store.registerClient({ ...clientMetadata, client_name: "Other Client" });

    expect(first.client_id).not.toBe(second.client_id);

    const afterRestart = new FileClientsStore(path);
    expect(afterRestart.getClient(first.client_id)).toBeDefined();
    expect(afterRestart.getClient(second.client_id)).toBeDefined();
    expect(afterRestart.size).toBe(2);
  });

  it("starts empty when the registry file does not exist", () => {
    const store = new FileClientsStore(join(dir, "does-not-exist.json"));
    expect(store.size).toBe(0);
  });

  it("ignores a corrupt registry instead of throwing", () => {
    writeFileSync(path, "{ not json", "utf8");
    const errorSpy = vi.spyOn(console, "error").mockImplementation(() => {
      // Silence the expected "registry is corrupt" warning.
    });

    const store = new FileClientsStore(path);

    expect(store.size).toBe(0);
    expect(errorSpy).toHaveBeenCalled();

    // A corrupt file must not block new registrations.
    const registered = store.registerClient(clientMetadata);
    expect(new FileClientsStore(path).getClient(registered.client_id)).toBeDefined();

    errorSpy.mockRestore();
  });

  it("skips entries without a client_id", () => {
    writeFileSync(path, JSON.stringify([{ client_name: "broken" }]), "utf8");
    expect(new FileClientsStore(path).size).toBe(0);
  });

  it("creates the data directory if it is missing", () => {
    const nested = join(dir, "nested", "deeper", "oauth-clients.json");
    const registered = new FileClientsStore(nested).registerClient(clientMetadata);

    const onDisk = JSON.parse(readFileSync(nested, "utf8"));
    expect(onDisk).toHaveLength(1);
    expect(onDisk[0].client_id).toBe(registered.client_id);
  });

  it("does not leave a temp file behind after writing", () => {
    const store = new FileClientsStore(path);
    store.registerClient(clientMetadata);

    expect(() => readFileSync(`${path}.tmp`, "utf8")).toThrow();
  });
});
