import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
const CACHE_PATH = join(homedir(), ".teams-mcp-token-cache.json");
/**
 * Custom file-based cache plugin for MSAL Node
 * Stores tokens (including refresh tokens) in a JSON file
 */
export const cachePlugin = {
    async beforeCacheAccess(cacheContext) {
        try {
            const data = await fs.readFile(CACHE_PATH, "utf8");
            cacheContext.tokenCache.deserialize(data);
        }
        catch (error) {
            // File doesn't exist or is invalid - start with empty cache
            if (error.code !== "ENOENT") {
                console.error("Warning: Could not read token cache:", error);
            }
        }
    },
    async afterCacheAccess(cacheContext) {
        if (cacheContext.cacheHasChanged) {
            try {
                const data = cacheContext.tokenCache.serialize();
                await fs.writeFile(CACHE_PATH, data, { encoding: "utf8", mode: 0o600 });
            }
            catch (error) {
                console.error("Warning: Could not write token cache:", error);
            }
        }
    },
};
export { CACHE_PATH };
//# sourceMappingURL=msal-cache.js.map