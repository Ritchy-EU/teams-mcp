import type { Client } from "@microsoft/microsoft-graph-client";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { GraphService } from "../../services/graph.js";
import type { FileUploadResult } from "../../utils/file-upload.js";
import { registerChatTools } from "../chats.js";

// Mock file-upload module
vi.mock("../../utils/file-upload.js", async () => {
  const actual = (await vi.importActual("../../utils/file-upload.js")) as any;
  return {
    ...actual,
    uploadFileToChat: vi.fn(),
  };
});

// Mock the Graph service
const mockGraphService = {
  getClient: vi.fn(),
} as unknown as GraphService;

// Mock the MCP server
const mockServer = {
  tool: vi.fn(),
} as unknown as McpServer;

// Mock client responses
const mockClient = {
  api: vi.fn(),
} as unknown as Client;

describe("Chat Tools", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockGraphService.getClient = vi.fn().mockResolvedValue(mockClient);
  });

  describe("registerChatTools", () => {
    it("should register all chat tools", () => {
      registerChatTools(mockServer, mockGraphService, false);

      expect(mockServer.tool).toHaveBeenCalledTimes(21);
      expect(mockServer.tool).toHaveBeenCalledWith(
        "list_chats",
        expect.any(String),
        {},
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "get_chat_messages",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "send_chat_message",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "create_chat",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "rename_chat",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "update_chat_message",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "delete_chat_message",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "download_chat_attachment",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "set_chat_message_reaction",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "unset_chat_message_reaction",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "send_file_to_chat",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "list_chat_members",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "add_chat_member",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "remove_chat_member",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "leave_chat",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "pin_chat_message",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "unpin_chat_message",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "hide_chat",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "unhide_chat",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "get_attachment_download_url",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "create_file_upload_session",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
    });

    it("should register only read-only chat tools when readOnly is true", () => {
      registerChatTools(mockServer, mockGraphService, true);

      expect(mockServer.tool).toHaveBeenCalledTimes(5);
      expect(mockServer.tool).toHaveBeenCalledWith(
        "get_attachment_download_url",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "list_chat_members",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "list_chats",
        expect.any(String),
        {},
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "get_chat_messages",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "download_chat_attachment",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
    });
  });

  describe("list_chats", () => {
    let listChatsHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi.mocked(mockServer.tool).mock.calls.find(([name]) => name === "list_chats");
      listChatsHandler = call?.[3] as unknown as (args: any) => Promise<any>;
    });

    it("should return chat list successfully", async () => {
      const mockChats = [
        {
          id: "chat1",
          topic: "Test Chat 1",
          chatType: "group",
          members: [{ displayName: "user1" }, { displayName: "user2" }],
        },
        {
          id: "chat2",
          topic: null,
          chatType: "oneOnOne",
          members: [{ displayName: "user1" }],
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockChats }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listChatsHandler();

      expect(mockClient.api).toHaveBeenCalledWith("/me/chats?$expand=members");
      expect(result.content[0].type).toBe("text");

      const parsedText = JSON.parse(result.content[0].text);
      expect(parsedText).toHaveLength(2);
      expect(parsedText[0]).toEqual({
        id: "chat1",
        topic: "Test Chat 1",
        chatType: "group",
        members: "user1, user2",
      });
      expect(parsedText[1]).toEqual({
        id: "chat2",
        topic: "No topic",
        chatType: "oneOnOne",
        members: "user1",
      });
    });

    it("should handle no chats found", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: [] }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listChatsHandler();

      expect(result.content[0].text).toBe("No chats found.");
    });

    it("should handle null response", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue(null),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listChatsHandler();

      expect(result.content[0].text).toBe("No chats found.");
    });

    it("should handle errors gracefully", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue(new Error("API Error")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listChatsHandler();

      expect(result.content[0].text).toBe("❌ Error: API Error");
    });

    it("should handle unknown errors", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue("Unknown error"),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listChatsHandler();

      expect(result.content[0].text).toBe("❌ Error: Unknown error occurred");
    });
  });

  describe("get_chat_messages", () => {
    let getChatMessagesHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "get_chat_messages");
      getChatMessagesHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should get chat messages with default parameters", async () => {
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "Hello world" },
          from: { user: { displayName: "John Doe" } },
          createdDateTime: "2023-01-01T10:00:00Z",
        },
        {
          id: "msg2",
          body: { content: "How are you?" },
          from: { user: { displayName: "Jane Smith" } },
          createdDateTime: "2023-01-01T11:00:00Z",
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({
        chatId: "chat123",
      });

      expect(mockClient.api).toHaveBeenCalledWith(
        "/me/chats/chat123/messages?$top=20&$orderby=createdDateTime desc"
      );

      const parsedResponse = JSON.parse(result.content[0].text);
      expect(parsedResponse.messages).toHaveLength(2);
      expect(parsedResponse.filteringMethod).toBe("server-side");
      expect(parsedResponse.totalReturned).toBe(2);
    });

    it("should apply all filtering options", async () => {
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "Hello" },
          from: { user: { displayName: "John" } },
          createdDateTime: "2023-01-01T10:00:00Z",
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const _result = await getChatMessagesHandler({
        chatId: "chat123",
        limit: 10,
        fromUser: "user123",
        orderBy: "lastModifiedDateTime",
        descending: true, // Changed to true since ascending is not supported
      });

      expect(mockClient.api).toHaveBeenCalledWith(
        "/me/chats/chat123/messages?$top=10&$orderby=lastModifiedDateTime desc&$filter=from/user/id eq 'user123'"
      );
    });

    it("should reject ascending order for datetime fields", async () => {
      const result = await getChatMessagesHandler({
        chatId: "chat123",
        orderBy: "lastModifiedDateTime",
        descending: false,
      });

      expect(result.content[0].text).toBe(
        "❌ Error: QueryOptions to order by 'LastModifiedDateTime' in 'Ascending' direction is not supported."
      );
    });

    it("should apply client-side date filtering", async () => {
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "Old message" },
          from: { user: { displayName: "John" } },
          createdDateTime: "2023-01-01T08:00:00Z", // Should be filtered out
        },
        {
          id: "msg2",
          body: { content: "New message" },
          from: { user: { displayName: "Jane" } },
          createdDateTime: "2023-01-01T12:00:00Z", // Should be included
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({
        chatId: "chat123",
        since: "2023-01-01T10:00:00Z",
        until: "2023-01-01T15:00:00Z",
      });

      const parsedResponse = JSON.parse(result.content[0].text);
      expect(parsedResponse.messages).toHaveLength(1);
      expect(parsedResponse.messages[0].content).toBe("New message");
      expect(parsedResponse.filteringMethod).toBe("client-side");
    });

    it("should handle messages without createdDateTime in date filtering", async () => {
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "Message without date" },
          from: { user: { displayName: "John" } },
          createdDateTime: null,
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({
        chatId: "chat123",
        since: "2023-01-01T10:00:00Z",
      });

      const parsedResponse = JSON.parse(result.content[0].text);
      expect(parsedResponse.messages).toHaveLength(1); // Should be included
    });

    it("should handle no messages found", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: [] }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe(
        "No messages found in this chat with the specified filters."
      );
    });

    it("should handle errors", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue(new Error("Chat not found")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe("❌ Error: Chat not found");
    });

    describe("pagination", () => {
      it("should fetch single page when fetchAll is false", async () => {
        const mockMessages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const mockApiChain = {
          get: vi.fn().mockResolvedValue({
            value: mockMessages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage",
          }),
        };
        mockClient.api = vi.fn().mockReturnValue(mockApiChain);

        const result = await getChatMessagesHandler({
          chatId: "chat123",
          limit: 100,
          fetchAll: false,
        });

        // Should only call API once
        expect(mockClient.api).toHaveBeenCalledTimes(1);

        const parsedResponse = JSON.parse(result.content[0].text);
        expect(parsedResponse.messages).toHaveLength(50);
      });

      it("should fetch multiple pages when fetchAll is true", async () => {
        const page1Messages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const page2Messages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i + 50}`,
          body: { content: `Message ${i + 50}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - (i + 50) * 1000).toISOString(),
        }));

        const page3Messages = Array.from({ length: 30 }, (_, i) => ({
          id: `msg${i + 100}`,
          body: { content: `Message ${i + 100}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - (i + 100) * 1000).toISOString(),
        }));

        const mockApiChain1 = {
          get: vi.fn().mockResolvedValue({
            value: page1Messages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage2",
          }),
        };

        const mockApiChain2 = {
          get: vi.fn().mockResolvedValue({
            value: page2Messages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage3",
          }),
        };

        const mockApiChain3 = {
          get: vi.fn().mockResolvedValue({
            value: page3Messages,
            "@odata.nextLink": undefined, // No more pages
          }),
        };

        mockClient.api = vi
          .fn()
          .mockReturnValueOnce(mockApiChain1)
          .mockReturnValueOnce(mockApiChain2)
          .mockReturnValueOnce(mockApiChain3);

        const result = await getChatMessagesHandler({
          chatId: "chat123",
          limit: 200,
          fetchAll: true,
        });

        // Should call API three times (initial + 2 pagination calls)
        expect(mockClient.api).toHaveBeenCalledTimes(3);

        const parsedResponse = JSON.parse(result.content[0].text);
        expect(parsedResponse.messages).toHaveLength(130); // 50 + 50 + 30
      });

      it("should stop fetching when limit is reached", async () => {
        const page1Messages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const page2Messages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i + 50}`,
          body: { content: `Message ${i + 50}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - (i + 50) * 1000).toISOString(),
        }));

        const mockApiChain1 = {
          get: vi.fn().mockResolvedValue({
            value: page1Messages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage2",
          }),
        };

        const mockApiChain2 = {
          get: vi.fn().mockResolvedValue({
            value: page2Messages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage3",
          }),
        };

        mockClient.api = vi
          .fn()
          .mockReturnValueOnce(mockApiChain1)
          .mockReturnValueOnce(mockApiChain2);

        const result = await getChatMessagesHandler({
          chatId: "chat123",
          limit: 75,
          fetchAll: true,
        });

        // Should only call API twice because limit is reached
        expect(mockClient.api).toHaveBeenCalledTimes(2);

        const parsedResponse = JSON.parse(result.content[0].text);
        // Should be limited to 75 messages even though 100 were fetched
        expect(parsedResponse.messages).toHaveLength(75);
      });

      it("should stop pagination when no nextLink is present", async () => {
        const mockMessages = Array.from({ length: 30 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const mockApiChain = {
          get: vi.fn().mockResolvedValue({
            value: mockMessages,
            "@odata.nextLink": undefined, // No more pages
          }),
        };
        mockClient.api = vi.fn().mockReturnValue(mockApiChain);

        const result = await getChatMessagesHandler({
          chatId: "chat123",
          limit: 100,
          fetchAll: true,
        });

        // Should only call API once since there's no nextLink
        expect(mockClient.api).toHaveBeenCalledTimes(1);

        const parsedResponse = JSON.parse(result.content[0].text);
        expect(parsedResponse.messages).toHaveLength(30);
      });

      it("should handle pagination errors gracefully", async () => {
        const page1Messages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const mockApiChain1 = {
          get: vi.fn().mockResolvedValue({
            value: page1Messages,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/nextPage2",
          }),
        };

        const mockApiChain2 = {
          get: vi.fn().mockRejectedValue(new Error("Network error")),
        };

        mockClient.api = vi
          .fn()
          .mockReturnValueOnce(mockApiChain1)
          .mockReturnValueOnce(mockApiChain2);

        const result = await getChatMessagesHandler({
          chatId: "chat123",
          limit: 100,
          fetchAll: true,
        });

        // Should return the messages from the first page even though second page failed
        const parsedResponse = JSON.parse(result.content[0].text);
        expect(parsedResponse.messages).toHaveLength(50);
      });

      it("should use smaller page size when fetchAll is true", async () => {
        const mockMessages = Array.from({ length: 50 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const mockApiChain = {
          get: vi.fn().mockResolvedValue({
            value: mockMessages,
          }),
        };
        mockClient.api = vi.fn().mockReturnValue(mockApiChain);

        await getChatMessagesHandler({
          chatId: "chat123",
          limit: 100,
          fetchAll: true,
        });

        // Should use page size of 50 instead of the full limit
        expect(mockClient.api).toHaveBeenCalledWith(
          "/me/chats/chat123/messages?$top=50&$orderby=createdDateTime desc"
        );
      });

      it("should use limit as page size when fetchAll is false and limit is small", async () => {
        const mockMessages = Array.from({ length: 10 }, (_, i) => ({
          id: `msg${i}`,
          body: { content: `Message ${i}` },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date(Date.now() - i * 1000).toISOString(),
        }));

        const mockApiChain = {
          get: vi.fn().mockResolvedValue({
            value: mockMessages,
          }),
        };
        mockClient.api = vi.fn().mockReturnValue(mockApiChain);

        await getChatMessagesHandler({
          chatId: "chat123",
          limit: 10,
          fetchAll: false,
        });

        // Should use limit (10) as page size since it's smaller than 50
        expect(mockClient.api).toHaveBeenCalledWith(
          "/me/chats/chat123/messages?$top=10&$orderby=createdDateTime desc"
        );
      });
    });
  });

  describe("send_chat_message", () => {
    let sendChatMessageHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "send_chat_message");
      sendChatMessageHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should send message with default importance", async () => {
      const mockResponse = { id: "newmsg123" };

      const mockApiChain = {
        post: vi.fn().mockResolvedValue(mockResponse),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "Hello world!",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/me/chats/chat123/messages");
      expect(mockApiChain.post).toHaveBeenCalledWith({
        body: {
          content: "Hello world!",
          contentType: "text",
        },
        importance: "normal",
      });
      expect(result.content[0].text).toBe("✅ Message sent successfully. Message ID: newmsg123");
    });

    it("should send message with custom importance", async () => {
      const mockResponse = { id: "newmsg456" };

      const mockApiChain = {
        post: vi.fn().mockResolvedValue(mockResponse),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "Urgent message",
        importance: "urgent",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        body: {
          content: "Urgent message",
          contentType: "text",
        },
        importance: "urgent",
      });
      expect(result.content[0].text).toBe("✅ Message sent successfully. Message ID: newmsg456");
    });

    it("should handle send errors", async () => {
      const mockApiChain = {
        post: vi.fn().mockRejectedValue(new Error("Permission denied")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "Test message",
      });

      expect(result.content[0].text).toBe("❌ Failed to send message: Permission denied");
    });

    it("should send message with markdown format", async () => {
      const mockResponse = { id: "mdmsg123" };
      const mockApiChain = { post: vi.fn().mockResolvedValue(mockResponse) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "**Bold** _Italic_",
        format: "markdown",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        body: {
          content: expect.stringContaining("<strong>Bold</strong>"),
          contentType: "html",
        },
        importance: "normal",
      });
      expect(result.content[0].text).toBe("✅ Message sent successfully. Message ID: mdmsg123");
    });

    it("should send message with text format (default)", async () => {
      const mockResponse = { id: "txtmsg123" };
      const mockApiChain = { post: vi.fn().mockResolvedValue(mockResponse) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "Plain text message",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        body: {
          content: "Plain text message",
          contentType: "text",
        },
        importance: "normal",
      });
      expect(result.content[0].text).toBe("✅ Message sent successfully. Message ID: txtmsg123");
    });

    it("should fallback to text for invalid format", async () => {
      const mockResponse = { id: "fallbackmsg123" };
      const mockApiChain = { post: vi.fn().mockResolvedValue(mockResponse) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendChatMessageHandler({
        chatId: "chat123",
        message: "Fallback message",
        format: "invalid-format",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        body: {
          content: "Fallback message",
          contentType: "text",
        },
        importance: "normal",
      });
      expect(result.content[0].text).toBe(
        "✅ Message sent successfully. Message ID: fallbackmsg123"
      );
    });
  });

  describe("update_chat_message", () => {
    let updateChatMessageHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "update_chat_message");
      updateChatMessageHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should update message with text content", async () => {
      const mockApiChain = {
        patch: vi.fn().mockResolvedValue(undefined),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await updateChatMessageHandler({
        chatId: "chat123",
        messageId: "msg456",
        message: "Updated text",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/me/chats/chat123/messages/msg456");
      expect(mockApiChain.patch).toHaveBeenCalledWith({
        body: {
          content: "Updated text",
          contentType: "text",
        },
      });
      expect(result.content[0].text).toBe("✅ Message updated successfully. Message ID: msg456");
    });

    it("should update message with explicit importance", async () => {
      const mockApiChain = {
        patch: vi.fn().mockResolvedValue(undefined),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      await updateChatMessageHandler({
        chatId: "chat123",
        messageId: "msg456",
        message: "Urgent update",
        importance: "urgent",
      });

      expect(mockApiChain.patch).toHaveBeenCalledWith({
        body: {
          content: "Urgent update",
          contentType: "text",
        },
        importance: "urgent",
      });
    });

    it("should update message with markdown format", async () => {
      const mockApiChain = {
        patch: vi.fn().mockResolvedValue(undefined),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await updateChatMessageHandler({
        chatId: "chat123",
        messageId: "msg456",
        message: "**Bold** _Italic_",
        format: "markdown",
      });

      expect(mockApiChain.patch).toHaveBeenCalledWith({
        body: {
          content: expect.stringContaining("<strong>Bold</strong>"),
          contentType: "html",
        },
      });
      expect(result.content[0].text).toContain("✅ Message updated successfully");
    });

    it("should update message with mentions", async () => {
      const mockPatchChain = {
        patch: vi.fn().mockResolvedValue(undefined),
      };
      const mockUserChain = {
        select: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ displayName: "John Doe" }),
        }),
      };

      mockClient.api = vi.fn().mockImplementation((url: string) => {
        if (url.startsWith("/users/")) {
          return mockUserChain;
        }
        return mockPatchChain;
      });

      const result = await updateChatMessageHandler({
        chatId: "chat123",
        messageId: "msg456",
        message: "Hello @johndoe!",
        mentions: [{ mention: "@johndoe", userId: "user-id-1" }],
      });

      expect(mockClient.api).toHaveBeenCalledWith("/users/user-id-1");
      expect(mockUserChain.select).toHaveBeenCalledWith("displayName");
      expect(mockPatchChain.patch).toHaveBeenCalledWith(
        expect.objectContaining({
          body: expect.objectContaining({ contentType: "html" }),
          mentions: expect.any(Array),
        })
      );
      expect(result.content[0].text).toContain("✅ Message updated successfully");
      expect(result.content[0].text).toContain("Mentions:");
    });

    it("should fallback to mention text when user resolution fails", async () => {
      const mockPatchChain = {
        patch: vi.fn().mockResolvedValue(undefined),
      };
      const mockUserChain = {
        select: vi.fn().mockReturnValue({
          get: vi.fn().mockRejectedValue(new Error("User not found")),
        }),
      };

      mockClient.api = vi.fn().mockImplementation((url: string) => {
        if (url.startsWith("/users/")) {
          return mockUserChain;
        }
        return mockPatchChain;
      });

      const consoleWarnSpy = vi.spyOn(console, "warn").mockImplementation(() => {
        // suppress console.warn in test
      });

      const result = await updateChatMessageHandler({
        chatId: "chat123",
        messageId: "msg456",
        message: "Hello @unknown!",
        mentions: [{ mention: "@unknown", userId: "bad-id" }],
      });

      expect(consoleWarnSpy).toHaveBeenCalledWith(
        expect.stringContaining("Could not resolve user bad-id")
      );
      expect(result.content[0].text).toContain("✅ Message updated successfully");
      consoleWarnSpy.mockRestore();
    });

    it("should handle update errors", async () => {
      const mockApiChain = {
        patch: vi.fn().mockRejectedValue(new Error("Permission denied")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await updateChatMessageHandler({
        chatId: "chat123",
        messageId: "msg456",
        message: "Updated text",
      });

      expect(result.content[0].text).toBe("❌ Failed to update message: Permission denied");
      expect(result.isError).toBe(true);
    });
  });

  describe("delete_chat_message", () => {
    let deleteChatMessageHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "delete_chat_message");
      deleteChatMessageHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should soft delete a chat message", async () => {
      const mockMeChain = {
        get: vi.fn().mockResolvedValue({ id: "current-user-id" }),
      };
      const mockDeleteChain = {
        post: vi.fn().mockResolvedValue(undefined),
      };

      mockClient.api = vi.fn().mockImplementation((url: string) => {
        if (url === "/me") {
          return mockMeChain;
        }
        return mockDeleteChain;
      });

      const result = await deleteChatMessageHandler({
        chatId: "chat123",
        messageId: "msg456",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/me");
      expect(mockClient.api).toHaveBeenCalledWith(
        "/users/current-user-id/chats/chat123/messages/msg456/softDelete"
      );
      expect(mockDeleteChain.post).toHaveBeenCalledWith({});
      expect(result.content[0].text).toBe("✅ Message deleted successfully. Message ID: msg456");
    });

    it("should handle delete errors", async () => {
      const mockMeChain = {
        get: vi.fn().mockResolvedValue({ id: "current-user-id" }),
      };
      const mockDeleteChain = {
        post: vi.fn().mockRejectedValue(new Error("Forbidden")),
      };

      mockClient.api = vi.fn().mockImplementation((url: string) => {
        if (url === "/me") {
          return mockMeChain;
        }
        return mockDeleteChain;
      });

      const result = await deleteChatMessageHandler({
        chatId: "chat123",
        messageId: "msg456",
      });

      expect(result.content[0].text).toBe("❌ Failed to delete message: Forbidden");
      expect(result.isError).toBe(true);
    });
  });

  describe("create_chat", () => {
    let createChatHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi.mocked(mockServer.tool).mock.calls.find(([name]) => name === "create_chat");
      createChatHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should create one-on-one chat", async () => {
      const mockMe = { id: "currentuser123" };
      const mockUser = { id: "otheruser456" };
      const mockNewChat = { id: "newchat789" };

      const mockApiChain = {
        get: vi
          .fn()
          .mockResolvedValueOnce(mockMe) // /me call
          .mockResolvedValueOnce(mockUser), // /users/email call
        post: vi.fn().mockResolvedValue(mockNewChat),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await createChatHandler({
        userEmails: ["other@example.com"],
      });

      expect(mockClient.api).toHaveBeenCalledWith("/me");
      expect(mockClient.api).toHaveBeenCalledWith("/users/other@example.com");
      expect(mockClient.api).toHaveBeenCalledWith("/chats");

      expect(mockApiChain.post).toHaveBeenCalledWith({
        chatType: "oneOnOne",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "currentuser123" },
            roles: ["owner"],
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "otheruser456" },
            roles: ["owner"],
          },
        ],
      });

      expect(result.content[0].text).toBe("✅ Chat created successfully. Chat ID: newchat789");
    });

    it("should create group chat with topic", async () => {
      const mockMe = { id: "currentuser123" };
      const mockUser1 = { id: "user1" };
      const mockUser2 = { id: "user2" };
      const mockNewChat = { id: "groupchat123" };

      const mockApiChain = {
        get: vi
          .fn()
          .mockResolvedValueOnce(mockMe) // /me call
          .mockResolvedValueOnce(mockUser1) // first user
          .mockResolvedValueOnce(mockUser2), // second user
        post: vi.fn().mockResolvedValue(mockNewChat),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const _result = await createChatHandler({
        userEmails: ["user1@example.com", "user2@example.com"],
        topic: "Project Discussion",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        chatType: "group",
        topic: "Project Discussion",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "currentuser123" },
            roles: ["owner"],
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "user1" },
            roles: ["owner"],
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "user2" },
            roles: ["owner"],
          },
        ],
      });
    });

    it("should ignore topic for one-on-one chats", async () => {
      const mockMe = { id: "currentuser123" };
      const mockUser = { id: "otheruser456" };
      const mockNewChat = { id: "newchat789" };

      const mockApiChain = {
        get: vi.fn().mockResolvedValueOnce(mockMe).mockResolvedValueOnce(mockUser),
        post: vi.fn().mockResolvedValue(mockNewChat),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const _result = await createChatHandler({
        userEmails: ["other@example.com"],
        topic: "This should be ignored",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        chatType: "oneOnOne",
        members: expect.any(Array),
      });

      // Should not include topic in the payload
      const postCall = mockApiChain.post.mock.calls[0][0];
      expect(postCall).not.toHaveProperty("topic");
    });

    it("should handle user lookup errors", async () => {
      const mockMe = { id: "currentuser123" };

      const mockApiChain = {
        get: vi
          .fn()
          .mockResolvedValueOnce(mockMe)
          .mockRejectedValueOnce(new Error("User not found")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await createChatHandler({
        userEmails: ["nonexistent@example.com"],
      });

      expect(result.content[0].text).toBe("❌ Error: User not found");
    });

    it("should handle chat creation errors", async () => {
      const mockMe = { id: "currentuser123" };
      const mockUser = { id: "otheruser456" };

      const mockApiChain = {
        get: vi.fn().mockResolvedValueOnce(mockMe).mockResolvedValueOnce(mockUser),
        post: vi.fn().mockRejectedValue(new Error("Failed to create chat")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await createChatHandler({
        userEmails: ["other@example.com"],
      });

      expect(result.content[0].text).toBe("❌ Error: Failed to create chat");
    });

    it("should exclude the creator's own email from userEmails", async () => {
      const mockMe = {
        id: "currentuser123",
        mail: "Me@Example.com",
        userPrincipalName: "me@example.com",
      };
      const mockUser = { id: "otheruser456" };
      const mockNewChat = { id: "newchat789" };

      const mockApiChain = {
        get: vi.fn().mockResolvedValueOnce(mockMe).mockResolvedValueOnce(mockUser),
        post: vi.fn().mockResolvedValue(mockNewChat),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await createChatHandler({
        userEmails: ["ME@example.com", "other@example.com"],
      });

      // Only /me, one user lookup and /chats — the creator is not looked up
      expect(mockClient.api).toHaveBeenCalledTimes(3);
      expect(mockClient.api).toHaveBeenCalledWith("/users/other@example.com");
      expect(mockApiChain.post).toHaveBeenCalledWith({
        chatType: "oneOnOne",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "currentuser123" },
            roles: ["owner"],
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: { id: "otheruser456" },
            roles: ["owner"],
          },
        ],
      });
      expect(result.content[0].text).toBe("✅ Chat created successfully. Chat ID: newchat789");
    });

    it("should dedupe repeated emails in userEmails", async () => {
      const mockMe = { id: "currentuser123" };
      const mockUser = { id: "otheruser456" };
      const mockNewChat = { id: "newchat789" };

      const mockApiChain = {
        get: vi.fn().mockResolvedValueOnce(mockMe).mockResolvedValueOnce(mockUser),
        post: vi.fn().mockResolvedValue(mockNewChat),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      await createChatHandler({
        userEmails: ["other@example.com", "Other@Example.com"],
      });

      // One lookup for the deduped address; resulting chat is 1:1
      expect(mockClient.api).toHaveBeenCalledTimes(3);
      const postCall = mockApiChain.post.mock.calls[0][0];
      expect(postCall.chatType).toBe("oneOnOne");
      expect(postCall.members).toHaveLength(2);
    });

    it("should fail when userEmails contains only the creator", async () => {
      const mockMe = {
        id: "currentuser123",
        mail: "me@example.com",
        userPrincipalName: "me@example.com",
      };
      const mockApiChain = {
        get: vi.fn().mockResolvedValueOnce(mockMe),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await createChatHandler({
        userEmails: ["me@example.com"],
      });

      expect(result.content[0].text).toBe(
        "❌ Error: No participants besides you — add at least one other user."
      );
      expect(mockClient.api).toHaveBeenCalledTimes(1);
    });
  });

  describe("rename_chat", () => {
    let renameChatHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService);
      const call = vi.mocked(mockServer.tool).mock.calls.find(([name]) => name === "rename_chat");
      renameChatHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should rename a group chat", async () => {
      const mockApiChain = {
        patch: vi.fn().mockResolvedValue({ id: "groupchat123", topic: "New Topic" }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await renameChatHandler({
        chatId: "groupchat123",
        topic: "New Topic",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/chats/groupchat123");
      expect(mockApiChain.patch).toHaveBeenCalledWith({ topic: "New Topic" });
      expect(result.content[0].text).toBe('✅ Chat renamed successfully. New topic: "New Topic"');
    });

    it("should reject topics containing ':'", async () => {
      const result = await renameChatHandler({
        chatId: "groupchat123",
        topic: "Project: Alpha",
      });

      expect(result.content[0].text).toBe(
        "❌ Error: Chat topic cannot contain ':' (not allowed by Microsoft Graph)."
      );
      expect(mockClient.api).not.toHaveBeenCalled();
    });

    it("should handle rename errors", async () => {
      const mockApiChain = {
        patch: vi.fn().mockRejectedValue(new Error("Cannot update topic for oneOnOne chat")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await renameChatHandler({
        chatId: "oneonone123",
        topic: "New Topic",
      });

      expect(result.content[0].text).toBe(
        "❌ Failed to rename chat: Cannot update topic for oneOnOne chat"
      );
    });
  });

  describe("get_chat_messages reactions", () => {
    let getChatMessagesHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "get_chat_messages");
      getChatMessagesHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should include reactions in message summaries", async () => {
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "Hello world" },
          from: { user: { displayName: "John Doe" } },
          createdDateTime: "2023-01-01T10:00:00Z",
          reactions: [
            {
              reactionType: "like",
              displayName: "Like",
              createdDateTime: "2023-01-01T10:01:00Z",
            },
            {
              reactionType: "heart",
              displayName: "Heart",
              createdDateTime: "2023-01-01T10:02:00Z",
            },
          ],
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({ chatId: "chat123" });
      const parsedResponse = JSON.parse(result.content[0].text);

      expect(parsedResponse.messages[0].reactions).toHaveLength(2);
      expect(parsedResponse.messages[0].reactions[0]).toEqual({
        reactionType: "like",
        displayName: "Like",
        createdDateTime: "2023-01-01T10:01:00Z",
      });
      expect(parsedResponse.messages[0].reactions[1]).toEqual({
        reactionType: "heart",
        displayName: "Heart",
        createdDateTime: "2023-01-01T10:02:00Z",
      });
    });

    it("should handle messages without reactions", async () => {
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "No reactions" },
          from: { user: { displayName: "John" } },
          createdDateTime: "2023-01-01T10:00:00Z",
        },
      ];

      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getChatMessagesHandler({ chatId: "chat123" });
      const parsedResponse = JSON.parse(result.content[0].text);

      expect(parsedResponse.messages[0].reactions).toBeUndefined();
    });
  });

  describe("set_chat_message_reaction", () => {
    let setReactionHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "set_chat_message_reaction");
      setReactionHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should set a reaction on a chat message", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue(undefined),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await setReactionHandler({
        chatId: "chat123",
        messageId: "msg456",
        reactionType: "like",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/chats/chat123/messages/msg456/setReaction");
      expect(mockApiChain.post).toHaveBeenCalledWith({ reactionType: "like" });
      expect(result.content[0].text).toBe("✅ Reaction like added to message msg456.");
    });

    it("should set a unicode emoji reaction", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue(undefined),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await setReactionHandler({
        chatId: "chat123",
        messageId: "msg456",
        reactionType: "👍",
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({ reactionType: "👍" });
      expect(result.content[0].text).toContain("👍");
    });

    it("should handle errors", async () => {
      const mockApiChain = {
        post: vi.fn().mockRejectedValue(new Error("Forbidden")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await setReactionHandler({
        chatId: "chat123",
        messageId: "msg456",
        reactionType: "like",
      });

      expect(result.content[0].text).toBe("❌ Failed to set reaction: Forbidden");
      expect(result.isError).toBe(true);
    });
  });

  describe("unset_chat_message_reaction", () => {
    let unsetReactionHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "unset_chat_message_reaction");
      unsetReactionHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should unset a reaction on a chat message", async () => {
      const mockApiChain = {
        post: vi.fn().mockResolvedValue(undefined),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await unsetReactionHandler({
        chatId: "chat123",
        messageId: "msg456",
        reactionType: "like",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/chats/chat123/messages/msg456/unsetReaction");
      expect(mockApiChain.post).toHaveBeenCalledWith({ reactionType: "like" });
      expect(result.content[0].text).toBe("✅ Reaction like removed from message msg456.");
    });

    it("should handle errors", async () => {
      const mockApiChain = {
        post: vi.fn().mockRejectedValue(new Error("Not found")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await unsetReactionHandler({
        chatId: "chat123",
        messageId: "msg456",
        reactionType: "like",
      });

      expect(result.content[0].text).toBe("❌ Failed to unset reaction: Not found");
      expect(result.isError).toBe(true);
    });
  });

  describe("send_file_to_chat", () => {
    let sendFileToChatHandler: (args?: any) => Promise<any>;

    beforeEach(async () => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "send_file_to_chat");
      sendFileToChatHandler = call?.[3] as unknown as (args: any) => Promise<any>;
    });

    it("should upload file and send message successfully", async () => {
      const { uploadFileToChat } = await import("../../utils/file-upload.js");

      const mockUploadResult: FileUploadResult = {
        webUrl: "https://onedrive.com/file.pdf",
        attachmentId: "AAAA-BBBB-CCCC",
        fileName: "report.pdf",
        fileSize: 2048,
        mimeType: "application/pdf",
      };
      vi.mocked(uploadFileToChat).mockResolvedValue(mockUploadResult);

      const mockApiChain = {
        post: vi.fn().mockResolvedValue({ id: "filemsg123" }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendFileToChatHandler({
        chatId: "chat123",
        filePath: "/tmp/report.pdf",
      });

      expect(result.content[0].text).toContain("✅ File sent successfully to chat.");
      expect(result.content[0].text).toContain("report.pdf");
      expect(result.content[0].text).toContain("Message ID: filemsg123");

      // Verify message payload has HTML body with attachment tag
      expect(mockApiChain.post).toHaveBeenCalledWith(
        expect.objectContaining({
          body: expect.objectContaining({
            content: expect.stringContaining('<attachment id="AAAA-BBBB-CCCC"></attachment>'),
            contentType: "html",
          }),
          attachments: [
            {
              id: "AAAA-BBBB-CCCC",
              contentType: "reference",
              contentUrl: "https://onedrive.com/file.pdf",
              name: "report.pdf",
            },
          ],
        })
      );
    });

    it("should include optional message text with HTML escaping", async () => {
      const { uploadFileToChat } = await import("../../utils/file-upload.js");

      const mockUploadResult: FileUploadResult = {
        webUrl: "https://onedrive.com/file.pdf",
        attachmentId: "AAAA-BBBB-CCCC",
        fileName: "report.pdf",
        fileSize: 1024,
        mimeType: "application/pdf",
      };
      vi.mocked(uploadFileToChat).mockResolvedValue(mockUploadResult);

      const mockApiChain = {
        post: vi.fn().mockResolvedValue({ id: "filemsg456" }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendFileToChatHandler({
        chatId: "chat123",
        filePath: "/tmp/report.pdf",
        message: "Check <this> file & report",
      });

      expect(result.content[0].text).toContain("✅ File sent successfully to chat.");

      // Plain text should be HTML-escaped in the body
      const postPayload = mockApiChain.post.mock.calls[0][0];
      expect(postPayload.body.content).toContain("Check &lt;this&gt; file &amp; report");
      expect(postPayload.body.content).toContain('<attachment id="AAAA-BBBB-CCCC"></attachment>');
    });

    it("should handle markdown format for message", async () => {
      const { uploadFileToChat } = await import("../../utils/file-upload.js");

      const mockUploadResult: FileUploadResult = {
        webUrl: "https://onedrive.com/file.pdf",
        attachmentId: "AAAA-BBBB-CCCC",
        fileName: "report.pdf",
        fileSize: 1024,
        mimeType: "application/pdf",
      };
      vi.mocked(uploadFileToChat).mockResolvedValue(mockUploadResult);

      const mockApiChain = {
        post: vi.fn().mockResolvedValue({ id: "filemsg789" }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendFileToChatHandler({
        chatId: "chat123",
        filePath: "/tmp/report.pdf",
        message: "**Bold** report",
        format: "markdown",
      });

      expect(result.content[0].text).toContain("✅ File sent successfully to chat.");

      const postPayload = mockApiChain.post.mock.calls[0][0];
      expect(postPayload.body.content).toContain("<strong>Bold</strong>");
      expect(postPayload.body.contentType).toBe("html");
    });

    it("should handle upload errors", async () => {
      const { uploadFileToChat } = await import("../../utils/file-upload.js");

      vi.mocked(uploadFileToChat).mockRejectedValue(new Error("File not found"));

      const result = await sendFileToChatHandler({
        chatId: "chat123",
        filePath: "/tmp/nonexistent.pdf",
      });

      expect(result.content[0].text).toBe("❌ Failed to send file: File not found");
      expect(result.isError).toBe(true);
    });

    it("should handle message send errors after successful upload", async () => {
      const { uploadFileToChat } = await import("../../utils/file-upload.js");

      const mockUploadResult: FileUploadResult = {
        webUrl: "https://onedrive.com/file.pdf",
        attachmentId: "AAAA-BBBB-CCCC",
        fileName: "report.pdf",
        fileSize: 1024,
        mimeType: "application/pdf",
      };
      vi.mocked(uploadFileToChat).mockResolvedValue(mockUploadResult);

      const mockApiChain = {
        post: vi.fn().mockRejectedValue(new Error("Permission denied")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await sendFileToChatHandler({
        chatId: "chat123",
        filePath: "/tmp/report.pdf",
      });

      expect(result.content[0].text).toBe("❌ Failed to send file: Permission denied");
      expect(result.isError).toBe(true);
    });
  });

  describe("list_chat_members", () => {
    let listMembersHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "list_chat_members");
      listMembersHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should list members with membership ids", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({
          value: [
            {
              id: "mem1",
              displayName: "User One",
              email: "one@example.com",
              userId: "u1",
              roles: ["owner"],
            },
            {
              id: "mem2",
              displayName: "User Two",
              email: "two@example.com",
              userId: "u2",
              roles: ["owner"],
            },
          ],
        }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listMembersHandler({ chatId: "chat123" });

      expect(mockClient.api).toHaveBeenCalledWith("/chats/chat123/members");
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
      expect(parsed[0]).toEqual({
        membershipId: "mem1",
        displayName: "User One",
        email: "one@example.com",
        userId: "u1",
        roles: ["owner"],
      });
    });

    it("should handle empty member list", async () => {
      const mockApiChain = { get: vi.fn().mockResolvedValue({ value: [] }) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listMembersHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe("No members found in this chat.");
    });

    it("should handle errors", async () => {
      const mockApiChain = { get: vi.fn().mockRejectedValue(new Error("Forbidden")) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await listMembersHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe("❌ Failed to list chat members: Forbidden");
      expect(result.isError).toBe(true);
    });
  });

  describe("add_chat_member", () => {
    let addMemberHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "add_chat_member");
      addMemberHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should add a member with full history by default", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ id: "user123" }),
        post: vi.fn().mockResolvedValue({}),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await addMemberHandler({
        chatId: "chat123",
        userEmail: "new@example.com",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/users/new@example.com");
      expect(mockClient.api).toHaveBeenCalledWith("/chats/chat123/members");
      expect(mockApiChain.post).toHaveBeenCalledWith({
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": "https://graph.microsoft.com/v1.0/users('user123')",
        visibleHistoryStartDateTime: "0001-01-01T00:00:00Z",
      });
      expect(result.content[0].text).toBe(
        "✅ new@example.com added to chat (full history visible)."
      );
    });

    it("should omit history when shareHistory is none", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ id: "user123" }),
        post: vi.fn().mockResolvedValue({}),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await addMemberHandler({
        chatId: "chat123",
        userEmail: "new@example.com",
        shareHistory: "none",
      });

      const payload = mockApiChain.post.mock.calls[0][0];
      expect(payload).not.toHaveProperty("visibleHistoryStartDateTime");
      expect(result.content[0].text).toBe("✅ new@example.com added to chat (no history visible).");
    });

    it("should use explicit shareHistorySince", async () => {
      const mockApiChain = {
        get: vi.fn().mockResolvedValue({ id: "user123" }),
        post: vi.fn().mockResolvedValue({}),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      await addMemberHandler({
        chatId: "chat123",
        userEmail: "new@example.com",
        shareHistorySince: "2026-08-01T00:00:00Z",
      });

      const payload = mockApiChain.post.mock.calls[0][0];
      expect(payload.visibleHistoryStartDateTime).toBe("2026-08-01T00:00:00Z");
    });

    it("should handle errors", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue(new Error("User not found")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await addMemberHandler({
        chatId: "chat123",
        userEmail: "ghost@example.com",
      });

      expect(result.content[0].text).toBe("❌ Failed to add chat member: User not found");
      expect(result.isError).toBe(true);
    });
  });

  describe("remove_chat_member", () => {
    let removeMemberHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "remove_chat_member");
      removeMemberHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should remove a member by email", async () => {
      const deleteMock = vi.fn().mockResolvedValue(undefined);
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/users/two@example.com") {
          return { get: vi.fn().mockResolvedValue({ id: "u2" }) };
        }
        if (path === "/chats/chat123/members") {
          return {
            get: vi.fn().mockResolvedValue({
              value: [
                { id: "mem1", userId: "u1" },
                { id: "mem2", userId: "u2" },
              ],
            }),
          };
        }
        if (path === "/chats/chat123/members/mem2") {
          return { delete: deleteMock };
        }
        return { get: vi.fn(), delete: vi.fn() };
      });

      const result = await removeMemberHandler({
        chatId: "chat123",
        userEmail: "two@example.com",
      });

      expect(deleteMock).toHaveBeenCalled();
      expect(result.content[0].text).toBe("✅ two@example.com removed from chat.");
    });

    it("should report when the user is not a member", async () => {
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/users/ghost@example.com") {
          return { get: vi.fn().mockResolvedValue({ id: "ghost" }) };
        }
        if (path === "/chats/chat123/members") {
          return {
            get: vi.fn().mockResolvedValue({ value: [{ id: "mem1", userId: "u1" }] }),
          };
        }
        return { get: vi.fn(), delete: vi.fn() };
      });

      const result = await removeMemberHandler({
        chatId: "chat123",
        userEmail: "ghost@example.com",
      });

      expect(result.content[0].text).toBe("❌ ghost@example.com is not a member of this chat.");
      expect(result.isError).toBe(true);
    });

    it("should handle errors", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue(new Error("Forbidden")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await removeMemberHandler({
        chatId: "chat123",
        userEmail: "two@example.com",
      });

      expect(result.content[0].text).toBe("❌ Failed to remove chat member: Forbidden");
      expect(result.isError).toBe(true);
    });
  });

  describe("leave_chat", () => {
    let leaveChatHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi.mocked(mockServer.tool).mock.calls.find(([name]) => name === "leave_chat");
      leaveChatHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should remove own membership", async () => {
      const deleteMock = vi.fn().mockResolvedValue(undefined);
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/me") {
          return { get: vi.fn().mockResolvedValue({ id: "me1" }) };
        }
        if (path === "/chats/chat123/members") {
          return {
            get: vi.fn().mockResolvedValue({
              value: [
                { id: "memMe", userId: "me1" },
                { id: "mem2", userId: "u2" },
              ],
            }),
          };
        }
        if (path === "/chats/chat123/members/memMe") {
          return { delete: deleteMock };
        }
        return { get: vi.fn(), delete: vi.fn() };
      });

      const result = await leaveChatHandler({ chatId: "chat123" });

      expect(deleteMock).toHaveBeenCalled();
      expect(result.content[0].text).toBe("✅ You left the chat.");
    });

    it("should report when not a member", async () => {
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/me") {
          return { get: vi.fn().mockResolvedValue({ id: "me1" }) };
        }
        if (path === "/chats/chat123/members") {
          return {
            get: vi.fn().mockResolvedValue({ value: [{ id: "mem2", userId: "u2" }] }),
          };
        }
        return { get: vi.fn(), delete: vi.fn() };
      });

      const result = await leaveChatHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe("❌ You are not a member of this chat.");
      expect(result.isError).toBe(true);
    });

    it("should handle errors", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue(new Error("Network error")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await leaveChatHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe("❌ Failed to leave chat: Network error");
      expect(result.isError).toBe(true);
    });
  });

  describe("pin_chat_message", () => {
    let pinHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "pin_chat_message");
      pinHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should pin a message", async () => {
      const mockApiChain = { post: vi.fn().mockResolvedValue({ id: "msg1" }) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await pinHandler({ chatId: "chat123", messageId: "msg1" });

      expect(mockClient.api).toHaveBeenCalledWith("/chats/chat123/pinnedMessages");
      expect(mockApiChain.post).toHaveBeenCalledWith({
        "message@odata.bind": "https://graph.microsoft.com/v1.0/chats/chat123/messages/msg1",
      });
      expect(result.content[0].text).toBe("✅ Message msg1 pinned.");
    });

    it("should handle errors", async () => {
      const mockApiChain = { post: vi.fn().mockRejectedValue(new Error("Forbidden")) };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await pinHandler({ chatId: "chat123", messageId: "msg1" });

      expect(result.content[0].text).toBe("❌ Failed to pin message: Forbidden");
      expect(result.isError).toBe(true);
    });
  });

  describe("unpin_chat_message", () => {
    let unpinHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "unpin_chat_message");
      unpinHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should unpin a message", async () => {
      const deleteMock = vi.fn().mockResolvedValue(undefined);
      mockClient.api = vi.fn().mockReturnValue({ delete: deleteMock });

      const result = await unpinHandler({ chatId: "chat123", messageId: "msg1" });

      expect(mockClient.api).toHaveBeenCalledWith("/chats/chat123/pinnedMessages/msg1");
      expect(deleteMock).toHaveBeenCalled();
      expect(result.content[0].text).toBe("✅ Message msg1 unpinned.");
    });

    it("should handle errors", async () => {
      mockClient.api = vi.fn().mockReturnValue({
        delete: vi.fn().mockRejectedValue(new Error("Not found")),
      });

      const result = await unpinHandler({ chatId: "chat123", messageId: "msg1" });

      expect(result.content[0].text).toBe("❌ Failed to unpin message: Not found");
      expect(result.isError).toBe(true);
    });
  });

  describe("hide_chat", () => {
    let hideHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi.mocked(mockServer.tool).mock.calls.find(([name]) => name === "hide_chat");
      hideHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should hide the chat for the current user", async () => {
      const postMock = vi.fn().mockResolvedValue(undefined);
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/me") {
          return { get: vi.fn().mockResolvedValue({ id: "me1" }) };
        }
        if (path === "/chats/chat123/members") {
          return {
            get: vi.fn().mockResolvedValue({
              value: [{ id: "memMe", userId: "me1", tenantId: "tenant-1" }],
            }),
          };
        }
        if (path === "/chats/chat123/hideForUser") {
          return { post: postMock };
        }
        return { get: vi.fn(), post: vi.fn() };
      });

      const result = await hideHandler({ chatId: "chat123" });

      expect(postMock).toHaveBeenCalledWith({
        user: { id: "me1", tenantId: "tenant-1" },
      });
      expect(result.content[0].text).toBe(
        "✅ Chat hidden from your chat list. It reappears on new activity."
      );
    });

    it("should report when not a member", async () => {
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/me") {
          return { get: vi.fn().mockResolvedValue({ id: "me1" }) };
        }
        if (path === "/chats/chat123/members") {
          return {
            get: vi.fn().mockResolvedValue({ value: [{ id: "mem2", userId: "u2" }] }),
          };
        }
        return { get: vi.fn(), post: vi.fn() };
      });

      const result = await hideHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe("❌ You are not a member of this chat.");
      expect(result.isError).toBe(true);
    });

    it("should handle errors", async () => {
      mockClient.api = vi.fn().mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("Network error")),
      });

      const result = await hideHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe("❌ Failed to hide chat: Network error");
      expect(result.isError).toBe(true);
    });
  });

  describe("unhide_chat", () => {
    let unhideHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi.mocked(mockServer.tool).mock.calls.find(([name]) => name === "unhide_chat");
      unhideHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should unhide the chat for the current user", async () => {
      const postMock = vi.fn().mockResolvedValue(undefined);
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/me") {
          return { get: vi.fn().mockResolvedValue({ id: "me1" }) };
        }
        if (path === "/chats/chat123/members") {
          return {
            get: vi.fn().mockResolvedValue({
              value: [{ id: "memMe", userId: "me1", tenantId: "tenant-1" }],
            }),
          };
        }
        if (path === "/chats/chat123/unhideForUser") {
          return { post: postMock };
        }
        return { get: vi.fn(), post: vi.fn() };
      });

      const result = await unhideHandler({ chatId: "chat123" });

      expect(postMock).toHaveBeenCalledWith({
        user: { id: "me1", tenantId: "tenant-1" },
      });
      expect(result.content[0].text).toBe("✅ Chat is visible in your chat list again.");
    });

    it("should handle errors", async () => {
      mockClient.api = vi.fn().mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("Network error")),
      });

      const result = await unhideHandler({ chatId: "chat123" });

      expect(result.content[0].text).toBe("❌ Failed to unhide chat: Network error");
      expect(result.isError).toBe(true);
    });
  });

  describe("get_attachment_download_url", () => {
    let getUrlHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "get_attachment_download_url");
      getUrlHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should return pre-authenticated download URLs for file attachments", async () => {
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/me/chats/chat123/messages/msg1") {
          return {
            get: vi.fn().mockResolvedValue({
              id: "msg1",
              attachments: [
                {
                  id: "att1",
                  contentType: "reference",
                  contentUrl: "https://contoso.sharepoint.com/personal/u/Documents/report.pdf",
                  name: "report.pdf",
                },
              ],
            }),
          };
        }
        if (path.startsWith("/shares/u!")) {
          return {
            get: vi.fn().mockResolvedValue({
              id: "item1",
              size: 2048,
              "@microsoft.graph.downloadUrl": "https://download.sharepoint.com/tempauth123",
            }),
          };
        }
        return { get: vi.fn() };
      });

      const result = await getUrlHandler({ chatId: "chat123", messageId: "msg1" });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.attachments).toHaveLength(1);
      expect(parsed.attachments[0]).toEqual({
        name: "report.pdf",
        size: 2048,
        downloadUrl: "https://download.sharepoint.com/tempauth123",
      });
    });

    it("should report when the message has no file attachments", async () => {
      mockClient.api = vi.fn().mockReturnValue({
        get: vi.fn().mockResolvedValue({ id: "msg1", attachments: [] }),
      });

      const result = await getUrlHandler({ chatId: "chat123", messageId: "msg1" });

      expect(result.content[0].text).toContain("No file attachments found");
    });

    it("should handle errors", async () => {
      mockClient.api = vi.fn().mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("Not found")),
      });

      const result = await getUrlHandler({ chatId: "chat123", messageId: "msg1" });

      expect(result.content[0].text).toBe("❌ Failed to get download URLs: Not found");
      expect(result.isError).toBe(true);
    });
  });

  describe("create_file_upload_session", () => {
    let createSessionHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "create_file_upload_session");
      createSessionHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should create an upload session in OneDrive for chat target", async () => {
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/me/drive") {
          return { get: vi.fn().mockResolvedValue({ id: "drive-1" }) };
        }
        if (path.includes("/createUploadSession")) {
          return {
            post: vi.fn().mockResolvedValue({
              uploadUrl: "https://up.1drv.com/session-abc",
              expirationDateTime: "2026-08-27T12:00:00Z",
            }),
          };
        }
        return { get: vi.fn(), post: vi.fn() };
      });

      const result = await createSessionHandler({ fileName: "big.zip", target: "chat" });

      const apiCalls = vi.mocked(mockClient.api).mock.calls.map((c: any[]) => c[0]);
      const sessionCall = apiCalls.find((p: string) => p.includes("/createUploadSession"));
      expect(sessionCall).toContain("/drives/drive-1/items/root:");
      expect(sessionCall).toContain(encodeURIComponent("Microsoft Teams Chat Files"));

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.uploadUrl).toBe("https://up.1drv.com/session-abc");
      expect(parsed.expiresAt).toBe("2026-08-27T12:00:00Z");
      expect(parsed.rules.chunkSizeMultipleOfBytes).toBe(327680);
      expect(parsed.examples.bash).toContain("curl");
    });

    it("should create an upload session in the channel SharePoint folder", async () => {
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/teams/team1/channels/chan1/filesFolder") {
          return {
            get: vi.fn().mockResolvedValue({
              id: "folder-1",
              parentReference: { driveId: "drive-9" },
            }),
          };
        }
        if (path.includes("/createUploadSession")) {
          return {
            post: vi.fn().mockResolvedValue({ uploadUrl: "https://up.spo.com/session-xyz" }),
          };
        }
        return { get: vi.fn(), post: vi.fn() };
      });

      const result = await createSessionHandler({
        fileName: "report.pdf",
        target: "channel",
        teamId: "team1",
        channelId: "chan1",
      });

      const apiCalls = vi.mocked(mockClient.api).mock.calls.map((c: any[]) => c[0]);
      const sessionCall = apiCalls.find((p: string) => p.includes("/createUploadSession"));
      expect(sessionCall).toContain("/drives/drive-9/items/folder-1:");

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.uploadUrl).toBe("https://up.spo.com/session-xyz");
    });

    it("should require teamId and channelId for channel target", async () => {
      const result = await createSessionHandler({ fileName: "a.txt", target: "channel" });

      expect(result.content[0].text).toBe(
        "❌ teamId and channelId are required when target is 'channel'."
      );
      expect(result.isError).toBe(true);
    });

    it("should handle errors", async () => {
      mockClient.api = vi.fn().mockReturnValue({
        get: vi.fn().mockRejectedValue(new Error("Drive unavailable")),
      });

      const result = await createSessionHandler({ fileName: "a.txt", target: "chat" });

      expect(result.content[0].text).toBe(
        "❌ Failed to create upload session: Drive unavailable"
      );
      expect(result.isError).toBe(true);
    });
  });

  describe("send_file_to_chat with driveItemId", () => {
    let sendFileHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerChatTools(mockServer, mockGraphService, false);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "send_file_to_chat");
      sendFileHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should send an already-uploaded drive item without re-uploading", async () => {
      const postMessageMock = vi.fn().mockResolvedValue({ id: "filemsg1" });
      mockClient.api = vi.fn().mockImplementation((path: string) => {
        if (path === "/me/drive") {
          return { get: vi.fn().mockResolvedValue({ id: "drive-1" }) };
        }
        if (path === "/drives/drive-1/items/item-42") {
          return {
            get: vi.fn().mockResolvedValue({
              id: "item-42",
              name: "big.zip",
              size: 123456,
              webUrl: "https://onedrive.com/big.zip",
              eTag: '"{AAAA-BBBB},1"',
            }),
          };
        }
        if (path === "/drives/drive-1/items/item-42/createLink") {
          return {
            post: vi.fn().mockResolvedValue({ link: { webUrl: "https://share.link/xyz" } }),
          };
        }
        if (path === "/me/chats/chat123/messages") {
          return { post: postMessageMock };
        }
        return { get: vi.fn(), post: vi.fn() };
      });

      const result = await sendFileHandler({ chatId: "chat123", driveItemId: "item-42" });

      expect(result.content[0].text).toContain("✅ File sent successfully to chat.");
      expect(result.content[0].text).toContain("big.zip");
      const payload = postMessageMock.mock.calls[0][0];
      expect(payload.attachments[0]).toEqual({
        id: "AAAA-BBBB",
        contentType: "reference",
        contentUrl: "https://share.link/xyz",
        name: "big.zip",
      });
    });

    it("should reject when both filePath and driveItemId are provided", async () => {
      const result = await sendFileHandler({
        chatId: "chat123",
        filePath: "/tmp/a.txt",
        driveItemId: "item-1",
      });

      expect(result.content[0].text).toBe("❌ Provide either filePath or driveItemId, not both.");
      expect(result.isError).toBe(true);
    });

    it("should reject when neither filePath nor driveItemId is provided", async () => {
      const result = await sendFileHandler({ chatId: "chat123" });

      expect(result.content[0].text).toContain("❌ Provide filePath");
      expect(result.isError).toBe(true);
    });
  });
});
