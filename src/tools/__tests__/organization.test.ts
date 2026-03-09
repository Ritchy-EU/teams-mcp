import { beforeEach, describe, expect, it, vi } from "vitest";
import { createMockMcpServer, mockUser } from "../../test-utils/setup.js";
import { registerOrganizationTools } from "../organization.js";

const mockManager = {
  id: "manager-id",
  displayName: "Manager User",
  userPrincipalName: "manager@example.com",
  mail: "manager@example.com",
  jobTitle: "Engineering Manager",
  department: "Engineering",
};

const mockDirector = {
  id: "director-id",
  displayName: "Director User",
  userPrincipalName: "director@example.com",
  mail: "director@example.com",
  jobTitle: "Director of Engineering",
  department: "Engineering",
};

const mockReport1 = {
  id: "report-1-id",
  displayName: "Report One",
  userPrincipalName: "report1@example.com",
  mail: "report1@example.com",
  jobTitle: "Software Engineer",
  department: "Engineering",
};

const mockReport2 = {
  id: "report-2-id",
  displayName: "Report Two",
  userPrincipalName: "report2@example.com",
  mail: "report2@example.com",
  jobTitle: "Software Engineer",
  department: "Engineering",
};

describe("Organization Tools", () => {
  let mockServer: any;
  let mockGraphService: any;
  let mockClient: any;

  beforeEach(() => {
    mockServer = createMockMcpServer();
    mockClient = {
      api: vi.fn().mockReturnValue({
        get: vi.fn(),
      }),
    };

    mockGraphService = {
      getClient: vi.fn().mockResolvedValue(mockClient),
    };

    vi.clearAllMocks();
  });

  describe("get_user_manager tool", () => {
    it("should register get_user_manager tool correctly", () => {
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_user_manager");
      expect(tool).toBeDefined();
    });

    it("should return manager for current user when no userId provided", async () => {
      mockClient.api.mockReturnValue({ get: vi.fn().mockResolvedValue(mockManager) });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_user_manager");
      const result = await tool.handler({});

      expect(mockClient.api).toHaveBeenCalledWith("/me/manager");
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.displayName).toBe("Manager User");
      expect(parsed.id).toBe("manager-id");
    });

    it("should return manager for a specific user", async () => {
      mockClient.api.mockReturnValue({ get: vi.fn().mockResolvedValue(mockManager) });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_user_manager");
      const result = await tool.handler({ userId: "test-user-id" });

      expect(mockClient.api).toHaveBeenCalledWith("/users/test-user-id/manager");
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.displayName).toBe("Manager User");
    });

    it("should handle user with no manager (top of org)", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("Resource 'manager' not found")),
      });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_user_manager");
      const result = await tool.handler({});

      expect(result.content[0].text).toContain("No manager found");
    });

    it("should handle API errors gracefully", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("API Error")),
      });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_user_manager");
      const result = await tool.handler({});

      expect(result.content[0].text).toBe("❌ Error: API Error");
    });

    it("should handle authentication errors", async () => {
      mockGraphService.getClient.mockRejectedValue(new Error("Not authenticated"));
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_user_manager");
      const result = await tool.handler({});

      expect(result.content[0].text).toContain("❌ Error: Not authenticated");
    });
  });

  describe("get_direct_reports tool", () => {
    it("should register get_direct_reports tool correctly", () => {
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_direct_reports");
      expect(tool).toBeDefined();
    });

    it("should return direct reports for current user", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue({ value: [mockReport1, mockReport2] }),
      });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_direct_reports");
      const result = await tool.handler({});

      expect(mockClient.api).toHaveBeenCalledWith("/me/directReports");
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
      expect(parsed[0].displayName).toBe("Report One");
      expect(parsed[1].displayName).toBe("Report Two");
    });

    it("should return direct reports for a specific user", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue({ value: [mockReport1] }),
      });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_direct_reports");
      const result = await tool.handler({ userId: "manager-id" });

      expect(mockClient.api).toHaveBeenCalledWith("/users/manager-id/directReports");
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(1);
    });

    it("should handle user with no direct reports", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue({ value: [] }),
      });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_direct_reports");
      const result = await tool.handler({});

      expect(result.content[0].text).toContain("No direct reports found");
    });

    it("should handle undefined value in response", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue({}),
      });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_direct_reports");
      const result = await tool.handler({});

      expect(result.content[0].text).toContain("No direct reports found");
    });

    it("should handle API errors gracefully", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("Forbidden")),
      });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_direct_reports");
      const result = await tool.handler({});

      expect(result.content[0].text).toBe("❌ Error: Forbidden");
    });
  });

  describe("get_manager_chain tool", () => {
    it("should register get_manager_chain tool correctly", () => {
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_manager_chain");
      expect(tool).toBeDefined();
    });

    it("should return full manager chain for current user", async () => {
      let callCount = 0;
      mockClient.api.mockImplementation(() => ({
        get: vi.fn().mockImplementation(() => {
          callCount++;
          if (callCount === 1) return Promise.resolve(mockManager);
          if (callCount === 2) return Promise.resolve(mockDirector);
          return Promise.reject(new Error("not found"));
        }),
      }));
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_manager_chain");
      const result = await tool.handler({});

      expect(mockClient.api).toHaveBeenCalledWith("/me/manager");
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
      expect(parsed[0].displayName).toBe("Manager User");
      expect(parsed[1].displayName).toBe("Director User");
    });

    it("should return manager chain for a specific user", async () => {
      let callCount = 0;
      mockClient.api.mockImplementation(() => ({
        get: vi.fn().mockImplementation(() => {
          callCount++;
          if (callCount === 1) return Promise.resolve(mockManager);
          return Promise.reject(new Error("not found"));
        }),
      }));
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_manager_chain");
      const result = await tool.handler({ userId: "test-user-id" });

      expect(mockClient.api).toHaveBeenCalledWith("/users/test-user-id/manager");
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(1);
      expect(parsed[0].displayName).toBe("Manager User");
    });

    it("should respect maxLevels parameter", async () => {
      let callCount = 0;
      mockClient.api.mockImplementation(() => ({
        get: vi.fn().mockImplementation(() => {
          callCount++;
          return Promise.resolve({
            ...mockManager,
            id: `manager-${callCount}`,
            displayName: `Manager Level ${callCount}`,
          });
        }),
      }));
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_manager_chain");
      const result = await tool.handler({ maxLevels: 2 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
    });

    it("should handle user at top of org (no managers)", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("not found")),
      });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_manager_chain");
      const result = await tool.handler({});

      expect(result.content[0].text).toContain("No managers found");
    });

    it("should handle authentication errors", async () => {
      mockGraphService.getClient.mockRejectedValue(new Error("Not authenticated"));
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_manager_chain");
      const result = await tool.handler({});

      expect(result.content[0].text).toContain("❌ Error: Not authenticated");
    });

    it("should stop when manager has no id", async () => {
      mockClient.api.mockReturnValue({
        get: vi.fn().mockResolvedValue({
          ...mockManager,
          id: undefined,
        }),
      });
      registerOrganizationTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_manager_chain");
      const result = await tool.handler({});

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(1);
    });
  });

  describe("authentication errors", () => {
    it("should handle authentication errors in all tools", async () => {
      const authError = new Error("Not authenticated");
      mockGraphService.getClient.mockRejectedValue(authError);
      registerOrganizationTools(mockServer, mockGraphService);

      const tools = ["get_user_manager", "get_direct_reports", "get_manager_chain"];

      for (const toolName of tools) {
        const tool = mockServer.getTool(toolName);
        const result = await tool.handler({});
        expect(result.content[0].text).toContain("❌ Error: Not authenticated");
      }
    });
  });
});
