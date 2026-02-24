import type { Channel, Chat, ChatMessage, ConversationMember, Team, User } from "../types/graph.js";
export declare const mockUser: User;
export declare const mockTeam: Team;
export declare const mockChannel: Channel;
export declare const mockChat: Chat;
export declare const mockChatMessage: ChatMessage;
export declare const mockConversationMember: ConversationMember;
export declare const graphApiHandlers: import("msw").HttpHandler[];
export declare const server: import("msw/node").SetupServerApi;
export declare function createMockGraphService(): {
    getInstance: import("vitest").Mock<import("@vitest/spy").Procedure>;
    getAuthStatus: import("vitest").Mock<import("@vitest/spy").Procedure>;
    getClient: import("vitest").Mock<import("@vitest/spy").Procedure>;
    isAuthenticated: import("vitest").Mock<import("@vitest/spy").Procedure>;
};
export declare function createMockUnauthenticatedGraphService(): {
    getInstance: import("vitest").Mock<import("@vitest/spy").Procedure>;
    getAuthStatus: import("vitest").Mock<import("@vitest/spy").Procedure>;
    getClient: import("vitest").Mock<import("@vitest/spy").Procedure>;
    isAuthenticated: import("vitest").Mock<import("@vitest/spy").Procedure>;
};
export declare function createMockMcpServer(): {
    tool: import("vitest").Mock<import("@vitest/spy").Procedure>;
    connect: import("vitest").Mock<import("@vitest/spy").Procedure>;
    getTool: (name: string) => any;
    getAllTools: () => any[];
};
export declare function testMcpTool(toolName: string, parameters: any, mockServer: any, expectedResult?: any): Promise<any>;
//# sourceMappingURL=setup.d.ts.map