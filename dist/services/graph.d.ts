import { Client } from "@microsoft/microsoft-graph-client";
export interface AuthStatus {
    isAuthenticated: boolean;
    userPrincipalName?: string | undefined;
    displayName?: string | undefined;
    expiresAt?: string | undefined;
}
export declare class GraphService {
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
    validateToken(token: string): string | undefined;
}
//# sourceMappingURL=graph.d.ts.map