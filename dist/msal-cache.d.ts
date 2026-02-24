import type { ICachePlugin } from "@azure/msal-node";
declare const CACHE_PATH: string;
/**
 * Custom file-based cache plugin for MSAL Node
 * Stores tokens (including refresh tokens) in a JSON file
 */
export declare const cachePlugin: ICachePlugin;
export { CACHE_PATH };
//# sourceMappingURL=msal-cache.d.ts.map