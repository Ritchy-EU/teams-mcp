import { Client } from "@microsoft/microsoft-graph-client";
export interface AuthStatus {
    isAuthenticated: boolean;
    userPrincipalName?: string | undefined;
    displayName?: string | undefined;
    expiresAt?: string | undefined;
}
/**
 * Common interface for Graph API access, used by both stdio (singleton) and HTTP (per-session) modes.
 */
export interface IGraphService {
    getAuthStatus(): Promise<AuthStatus>;
    getClient(): Promise<Client>;
    isAuthenticated(): boolean;
    startDeviceCodeAuth(): Promise<{
        verificationUri: string;
        userCode: string;
        expiresIn: number;
    }>;
    validateToken(token: string): string | undefined;
}
/**
 * Validate a Microsoft Graph JWT token (shared logic).
 */
export declare function validateGraphToken(token: string): string | undefined;
/**
 * Graph service for HTTP mode: uses a per-session token accessor instead of MSAL cache.
 */
export declare class SessionGraphService implements IGraphService {
    private tokenAccessor;
    private client;
    constructor(tokenAccessor: () => Promise<string>);
    getClient(): Promise<Client>;
    getAuthStatus(): Promise<AuthStatus>;
    isAuthenticated(): boolean;
    startDeviceCodeAuth(): Promise<never>;
    validateToken(token: string): string | undefined;
}
/**
 * Graph service for stdio mode: singleton with MSAL-backed token cache.
 */
export declare class GraphService implements IGraphService {
    private static instance;
    private client;
    private isInitialized;
    private tokenExpiresAt;
    private msalApp;
    private msalAccount;
    static getInstance(): GraphService;
    private initializeClient;
    private acquireToken;
    getAuthStatus(): Promise<AuthStatus>;
    getClient(): Promise<Client>;
    isAuthenticated(): boolean;
    startDeviceCodeAuth(): Promise<{
        verificationUri: string;
        userCode: string;
        expiresIn: number;
    }>;
    validateToken(token: string): string | undefined;
}
//# sourceMappingURL=graph.d.ts.map