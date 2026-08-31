import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { IGraphService } from "../services/graph.js";
import type {
  AadUserConversationMember,
  Chat,
  ChatMessage,
  ChatMessageReaction,
  ChatSummary,
  ConversationMember,
  CreateChatPayload,
  GraphApiResponse,
  MessageSummary,
  ReactionSummary,
  User,
} from "../types/graph.js";
import { extractAttachmentSummaries } from "../utils/attachments.js";
import {
  buildFileAttachment,
  createUploadSessionForChannel,
  createUploadSessionForChat,
  escapeHtml,
  type FileUploadResult,
  formatFileSize,
  resolveChatDriveItem,
  type UploadSessionInfo,
  uploadFileToChat,
} from "../utils/file-upload.js";
import { markdownToHtml } from "../utils/markdown.js";
import { processMentionsInHtml } from "../utils/users.js";

/** Attachments larger than this are refused for inline base64 download. */
const MAX_INLINE_ATTACHMENT_BYTES = 1024 * 1024; // 1 MiB

/** Encode a SharePoint/OneDrive URL for the Graph Shares API ("u!" base64url format). */
function encodeShareUrl(url: string): string {
  return `u!${Buffer.from(url)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "")}`;
}

/**
 * Registers all chat-related MCP tools on the given server.
 * Tools include: list_chats, get_chat_messages, send_chat_message,
 * create_chat, update_chat_message, and delete_chat_message.
 *
 * @param server - The MCP server instance to register tools on.
 * @param graphService - The Microsoft Graph service used for API calls.
 */
export function registerChatTools(
  server: McpServer,
  graphService: IGraphService,
  readOnly: boolean
) {
  // List user's chats
  server.tool(
    "list_chats",
    "List all recent chats (1:1 conversations and group chats) that the current user participates in. Returns chat topics, types, and participant information.",
    {},
    async () => {
      try {
        // Build query parameters
        const queryParams: string[] = ["$expand=members"];

        const queryString = queryParams.join("&");

        const client = await graphService.getClient();
        const response = (await client
          .api(`/me/chats?${queryString}`)
          .get()) as GraphApiResponse<Chat>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No chats found.",
              },
            ],
          };
        }

        const chatList: ChatSummary[] = response.value.map((chat: Chat) => ({
          id: chat.id,
          topic: chat.topic || "No topic",
          chatType: chat.chatType,
          members:
            chat.members?.map((member: ConversationMember) => member.displayName).join(", ") ||
            "No members",
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(chatList, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Get chat messages with pagination support
  server.tool(
    "get_chat_messages",
    "Retrieve recent messages from a specific chat conversation. Returns message content, sender information, and timestamps.",
    {
      chatId: z.string().describe("Chat ID (e.g. 19:meeting_Njhi..j@thread.v2"),
      limit: z
        .number()
        .min(1)
        .max(2000)
        .optional()
        .default(20)
        .describe("Number of messages to retrieve (default: 20, max: 2000)"),
      since: z.string().optional().describe("Get messages since this ISO datetime"),
      until: z.string().optional().describe("Get messages until this ISO datetime"),
      fromUser: z.string().optional().describe("Filter messages from specific user ID"),
      orderBy: z
        .enum(["createdDateTime", "lastModifiedDateTime"])
        .optional()
        .default("createdDateTime")
        .describe("Sort order"),
      descending: z
        .boolean()
        .optional()
        .default(true)
        .describe("Sort in descending order (newest first)"),
      fetchAll: z
        .boolean()
        .optional()
        .default(false)
        .describe(
          "Fetch all messages using pagination (up to limit). When true, follows @odata.nextLink to get more messages."
        ),
    },
    async ({ chatId, limit, since, until, fromUser, orderBy, descending, fetchAll }) => {
      try {
        const client = await graphService.getClient();

        // Apply defaults for parameters (in case Zod validation is bypassed)
        const effectiveLimit = limit ?? 20;
        const effectiveOrderBy = orderBy ?? "createdDateTime";
        const effectiveDescending = descending ?? true;
        const effectiveFetchAll = fetchAll ?? false;

        // Build query parameters - use smaller page size for pagination
        const pageSize = effectiveFetchAll ? 50 : Math.min(effectiveLimit, 50);
        const queryParams: string[] = [`$top=${pageSize}`];

        // Add ordering - Graph API only supports descending order for datetime fields in chat messages
        if (
          (effectiveOrderBy === "createdDateTime" || effectiveOrderBy === "lastModifiedDateTime") &&
          !effectiveDescending
        ) {
          return {
            content: [
              {
                type: "text",
                text: `❌ Error: QueryOptions to order by '${effectiveOrderBy === "createdDateTime" ? "CreatedDateTime" : "LastModifiedDateTime"}' in 'Ascending' direction is not supported.`,
              },
            ],
          };
        }

        const sortDirection = effectiveDescending ? "desc" : "asc";
        queryParams.push(`$orderby=${effectiveOrderBy} ${sortDirection}`);

        // Add filters (only user filter is supported reliably)
        const filters: string[] = [];
        if (fromUser) {
          filters.push(`from/user/id eq '${fromUser}'`);
        }

        if (filters.length > 0) {
          queryParams.push(`$filter=${filters.join(" and ")}`);
        }

        const queryString = queryParams.join("&");

        // Fetch messages with pagination support
        const allMessages: ChatMessage[] = [];
        let nextLink: string | undefined;
        let pageCount = 0;
        const maxPages = 100; // Safety limit to prevent infinite loops

        // First request
        let response = (await client
          .api(`/me/chats/${chatId}/messages?${queryString}`)
          .get()) as GraphApiResponse<ChatMessage>;

        if (response?.value) {
          allMessages.push(...response.value);
        }

        // Follow pagination if fetchAll is enabled
        if (effectiveFetchAll) {
          nextLink = response["@odata.nextLink"];

          while (nextLink && allMessages.length < effectiveLimit && pageCount < maxPages) {
            pageCount++;

            try {
              response = (await client.api(nextLink).get()) as GraphApiResponse<ChatMessage>;

              if (response?.value) {
                allMessages.push(...response.value);
              }

              nextLink = response["@odata.nextLink"];
            } catch (pageError) {
              console.error(`Error fetching page ${pageCount}:`, pageError);
              break;
            }
          }
        }

        if (allMessages.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: "No messages found in this chat with the specified filters.",
              },
            ],
          };
        }

        // Apply client-side date filtering since server-side filtering is not supported
        let filteredMessages = allMessages;

        if (since || until) {
          filteredMessages = allMessages.filter((message: ChatMessage) => {
            if (!message.createdDateTime) return true;

            const messageDate = new Date(message.createdDateTime);
            if (since) {
              const sinceDate = new Date(since);
              if (messageDate <= sinceDate) return false;
            }
            if (until) {
              const untilDate = new Date(until);
              if (messageDate >= untilDate) return false;
            }
            return true;
          });
        }

        // Apply limit after filtering
        const limitedMessages = filteredMessages.slice(0, effectiveLimit);

        const messageList: MessageSummary[] = limitedMessages.map((message: ChatMessage) => {
          const summary: MessageSummary = {
            id: message.id,
            content: message.body?.content,
            from: message.from?.user?.displayName,
            createdDateTime: message.createdDateTime,
          };

          // Include attachment metadata if present
          summary.attachments = extractAttachmentSummaries(message.attachments);

          // Include reactions if present
          if (message.reactions?.length) {
            summary.reactions = message.reactions.map(
              (r: ChatMessageReaction): ReactionSummary => ({
                reactionType: r.reactionType,
                displayName: r.displayName,
                createdDateTime: r.createdDateTime,
              })
            );
          }

          return summary;
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  filters: { since, until, fromUser },
                  filteringMethod: since || until ? "client-side" : "server-side",
                  paginationEnabled: fetchAll,
                  pagesRetrieved: pageCount + 1,
                  totalRetrieved: allMessages.length,
                  totalReturned: messageList.length,
                  hasMore: !!response["@odata.nextLink"] || filteredMessages.length > limit,
                  messages: messageList,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Download attachments and hosted content from chat messages
  server.tool(
    "download_chat_attachment",
    "Download files and images from a chat message. Handles both inline images (hosted content) and file attachments (OneDrive/SharePoint references). Use get_chat_messages first to see available attachments.",
    {
      chatId: z.string().describe("Chat ID"),
      messageId: z.string().describe("Message ID containing the attachment"),
      attachmentIndex: z
        .number()
        .min(0)
        .optional()
        .describe(
          "Index of a specific attachment to download (0-based). If not provided, downloads all attachments."
        ),
      savePath: z
        .string()
        .optional()
        .describe(
          "Optional file path to save the content. Supports UNC paths (e.g., \\\\wsl.localhost\\Ubuntu\\tmp\\file.png)."
        ),
    },
    async ({ chatId, messageId, attachmentIndex, savePath }) => {
      try {
        const client = await graphService.getClient();

        // Fetch the message to inspect attachments and body
        const message = (await client
          .api(`/me/chats/${chatId}/messages/${messageId}`)
          .get()) as ChatMessage;

        if (!message) {
          return {
            content: [{ type: "text", text: "❌ Message not found." }],
            isError: true,
          };
        }

        // Collect downloadable items
        interface DownloadItem {
          type: "hostedContent" | "fileReference";
          id: string;
          name: string;
          contentUrl?: string | null;
        }

        const items: DownloadItem[] = [];

        // 1. Extract hosted content IDs from message body HTML (inline images)
        const bodyContent = message.body?.content || "";
        const hostedContentRegex = /hostedContents\/([a-zA-Z0-9_=-]+)\/\$value|itemid="([^"]+)"/gi;
        let match: RegExpExecArray | null;

        // biome-ignore lint/suspicious/noAssignInExpressions: needed for regex extraction
        while ((match = hostedContentRegex.exec(bodyContent)) !== null) {
          const contentId = match[1] || match[2];
          if (contentId && !items.some((i) => i.id === contentId)) {
            items.push({
              type: "hostedContent",
              id: contentId,
              name: `hosted_image_${items.length}.png`,
            });
          }
        }

        // 2. Collect file reference attachments
        if (message.attachments) {
          for (const att of message.attachments) {
            if (att.contentType === "reference" && att.contentUrl) {
              items.push({
                type: "fileReference",
                id: att.id || `ref_${items.length}`,
                name: att.name || "unknown_file",
                contentUrl: att.contentUrl,
              });
            }
          }
        }

        if (items.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: "❌ No downloadable attachments found in this message.",
              },
            ],
            isError: true,
          };
        }

        // Filter to specific attachment if index provided
        const targetItems =
          attachmentIndex !== undefined ? [items[attachmentIndex]].filter(Boolean) : items;

        if (targetItems.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: `❌ Attachment index ${attachmentIndex} out of range. Message has ${items.length} attachment(s).`,
              },
            ],
            isError: true,
          };
        }

        // Download each item
        const results: Array<{
          index: number;
          type: string;
          name: string;
          size: number;
          base64Data?: string;
          savedTo?: string;
          error?: string;
        }> = [];

        for (const item of targetItems) {
          const itemIndex = items.indexOf(item);
          try {
            let buffer: Buffer;

            if (item.type === "hostedContent") {
              // Download hosted content (inline images)
              const response = await client
                .api(`/chats/${chatId}/messages/${messageId}/hostedContents/${item.id}/$value`)
                .responseType("arraybuffer" as any)
                .get();
              buffer = Buffer.from(response as ArrayBuffer);
            } else {
              // Download file reference via Shares API
              if (!item.contentUrl) {
                throw new Error("Attachment has no download URL (contentUrl missing)");
              }
              const encodedUrl = encodeShareUrl(item.contentUrl);

              const response = await client
                .api(`/shares/${encodedUrl}/driveItem/content`)
                .responseType("arraybuffer" as any)
                .get();
              buffer = Buffer.from(response as ArrayBuffer);
            }

            const result: {
              index: number;
              type: string;
              name: string;
              size: number;
              base64Data?: string;
              savedTo?: string;
            } = {
              index: itemIndex,
              type: item.type,
              name: item.name,
              size: buffer.length,
            };

            // Save to disk or return base64
            if (savePath) {
              const fs = await import("node:fs/promises");
              const path = await import("node:path");

              const normalizedPath = savePath.replace(/\\\\/g, "\\");
              const isUncPath =
                normalizedPath.startsWith("\\\\") || normalizedPath.startsWith("//");

              let finalPath = normalizedPath;
              if (targetItems.length > 1) {
                const ext = path.extname(normalizedPath);
                const base = ext ? normalizedPath.slice(0, -ext.length) : normalizedPath;
                finalPath = `${base}_${itemIndex}${ext || path.extname(item.name)}`;
              }

              const targetPath = isUncPath ? finalPath : path.resolve(finalPath);
              await fs.writeFile(targetPath, buffer);
              result.savedTo = targetPath;
            } else if (buffer.length > MAX_INLINE_ATTACHMENT_BYTES) {
              throw new Error(
                `File is too large for inline base64 (${buffer.length} bytes, limit ${MAX_INLINE_ATTACHMENT_BYTES}). ` +
                  "Use get_attachment_download_url to download it directly instead."
              );
            } else {
              result.base64Data = buffer.toString("base64");
            }

            results.push(result);
          } catch (downloadError) {
            const errorMsg =
              downloadError instanceof Error ? downloadError.message : "Unknown error";
            results.push({
              index: itemIndex,
              type: item.type,
              name: item.name,
              size: 0,
              error: errorMsg,
            });
          }
        }

        const successCount = results.filter((r) => !r.error).length;
        const errorCount = results.filter((r) => r.error).length;

        let summary = `📥 Downloaded ${successCount} of ${targetItems.length} attachment(s)`;
        if (errorCount > 0) {
          summary += ` (${errorCount} failed)`;
        }

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  summary,
                  messageId,
                  totalAttachments: items.length,
                  downloaded: targetItems.length,
                  successCount,
                  errorCount,
                  attachments: results,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // List members of a chat
  server.tool(
    "list_chat_members",
    "List all members of a chat with their membership IDs, names, emails and roles.",
    {
      chatId: z.string().describe("Chat ID"),
    },
    async ({ chatId }) => {
      try {
        const client = await graphService.getClient();

        const response = (await client
          .api(`/chats/${chatId}/members`)
          .get()) as GraphApiResponse<AadUserConversationMember>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text" as const,
                text: "No members found in this chat.",
              },
            ],
          };
        }

        const members = response.value.map((member) => ({
          membershipId: member.id,
          displayName: member.displayName,
          email: member.email,
          userId: member.userId,
          roles: member.roles,
        }));

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(members, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to list chat members: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Get short-lived direct download URLs for message file attachments
  server.tool(
    "get_attachment_download_url",
    "Get short-lived (~1 hour) pre-authenticated download URLs for the file attachments of a chat message. The URLs download directly from OneDrive/SharePoint without auth headers (curl/browser friendly) — preferred over download_chat_attachment for anything larger than ~1 MB. Inline images (hosted content) have no download URL; use download_chat_attachment for those.",
    {
      chatId: z.string().describe("Chat ID"),
      messageId: z.string().describe("Message ID containing the attachments"),
    },
    async ({ chatId, messageId }) => {
      try {
        const client = await graphService.getClient();

        const message = (await client
          .api(`/me/chats/${chatId}/messages/${messageId}`)
          .get()) as ChatMessage;

        const fileAttachments = (message.attachments ?? []).filter(
          (att) => att.contentType === "reference" && att.contentUrl
        );

        if (fileAttachments.length === 0) {
          return {
            content: [
              {
                type: "text" as const,
                text: "No file attachments found in this message. Inline images (hosted content) can be fetched with download_chat_attachment.",
              },
            ],
          };
        }

        const results: Array<Record<string, unknown>> = [];
        for (const att of fileAttachments) {
          const contentUrl = att.contentUrl;
          if (!contentUrl) {
            continue;
          }
          try {
            const item = (await client
              .api(`/shares/${encodeShareUrl(contentUrl)}/driveItem`)
              .get()) as Record<string, unknown>;
            results.push({
              name: att.name,
              size: item.size,
              downloadUrl: item["@microsoft.graph.downloadUrl"],
            });
          } catch (itemError: unknown) {
            results.push({
              name: att.name,
              error: itemError instanceof Error ? itemError.message : "Unknown error",
            });
          }
        }

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  note: "Download URLs are pre-authenticated and expire after about 1 hour.",
                  attachments: results,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to get download URLs: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Write tools below are skipped when the server runs in read-only mode.
  if (readOnly) return;

  // Send chat message
  server.tool(
    "send_chat_message",
    "Send a message to a specific chat conversation. Supports text and markdown formatting, mentions, and importance levels.",
    {
      chatId: z.string().describe("Chat ID"),
      message: z.string().describe("Message content"),
      importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
      format: z.enum(["text", "markdown"]).optional().describe("Message format (text or markdown)"),
      mentions: z
        .array(
          z.object({
            mention: z
              .string()
              .describe("The @mention text (e.g., 'john.doe' or 'john.doe@company.com')"),
            userId: z.string().describe("Azure AD User ID of the mentioned user"),
          })
        )
        .optional()
        .describe("Array of @mentions to include in the message"),
    },
    async ({ chatId, message, importance = "normal", format = "text", mentions }) => {
      try {
        const client = await graphService.getClient();

        // Process message content based on format
        let content: string;
        let contentType: "text" | "html";

        if (format === "markdown") {
          content = await markdownToHtml(message);
          contentType = "html";
        } else {
          content = message;
          contentType = "text";
        }

        // Process @mentions if provided
        const mentionMappings: Array<{ mention: string; userId: string; displayName: string }> = [];
        if (mentions && mentions.length > 0) {
          // Convert provided mentions to mappings with display names
          for (const mention of mentions) {
            try {
              // Get user info to get display name
              const userResponse = await client
                .api(`/users/${mention.userId}`)
                .select("displayName")
                .get();
              mentionMappings.push({
                mention: mention.mention,
                userId: mention.userId,
                displayName: userResponse.displayName || mention.mention,
              });
            } catch (_error) {
              console.warn(
                `Could not resolve user ${mention.userId}, using mention text as display name`
              );
              mentionMappings.push({
                mention: mention.mention,
                userId: mention.userId,
                displayName: mention.mention,
              });
            }
          }
        }

        // Process mentions in HTML content
        let finalMentions: Array<{
          id: number;
          mentionText: string;
          mentioned: { user: { id: string } };
        }> = [];
        if (mentionMappings.length > 0) {
          const result = processMentionsInHtml(content, mentionMappings);
          content = result.content;
          finalMentions = result.mentions;

          // Ensure we're using HTML content type when mentions are present
          contentType = "html";
        }

        // Build message payload
        const messagePayload: any = {
          body: {
            content,
            contentType,
          },
          importance,
        };

        if (finalMentions.length > 0) {
          messagePayload.mentions = finalMentions;
        }

        const result = (await client
          .api(`/me/chats/${chatId}/messages`)
          .post(messagePayload)) as ChatMessage;

        // Build success message
        const successText = `✅ Message sent successfully. Message ID: ${result.id}${
          finalMentions.length > 0
            ? `\n📱 Mentions: ${finalMentions.map((m) => m.mentionText).join(", ")}`
            : ""
        }`;

        return {
          content: [
            {
              type: "text" as const,
              text: successText,
            },
          ],
        };
      } catch (error: any) {
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to send message: ${error.message}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Create new chat (1:1 or group)
  server.tool(
    "create_chat",
    "Create a new chat conversation. Can be a 1:1 chat (with one other user) or a group chat (with multiple users). Group chats can optionally have a topic. The current user is added automatically and must not be included in userEmails.",
    {
      userEmails: z.array(z.string()).describe("Array of user email addresses to add to chat"),
      topic: z.string().optional().describe("Chat topic (for group chats)"),
    },
    async ({ userEmails, topic }) => {
      try {
        const client = await graphService.getClient();

        // Get current user ID
        const me = (await client.api("/me").get()) as User;

        // The creator is added automatically below; drop their own address and
        // any repeats from userEmails, or Graph rejects the request with
        // "Duplicate chat members is specified in the request body".
        const ownAddresses = [me?.mail, me?.userPrincipalName]
          .filter((a): a is string => !!a)
          .map((a) => a.toLowerCase());
        const uniqueEmails = [...new Set(userEmails.map((e) => e.toLowerCase()))].filter(
          (email) => !ownAddresses.includes(email)
        );

        if (uniqueEmails.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: "❌ Error: No participants besides you — add at least one other user.",
              },
            ],
          };
        }

        // Create members array
        const members: ConversationMember[] = [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: {
              id: me?.id,
            },
            roles: ["owner"],
          } as ConversationMember,
        ];

        // Add other users as members.
        // Graph only accepts the "owner" role when creating a chat — "member"
        // is rejected with "The passed-in role 'member' is not supported".
        for (const email of uniqueEmails) {
          const user = (await client.api(`/users/${email}`).get()) as User;
          members.push({
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            user: {
              id: user?.id,
            },
            roles: ["owner"],
          } as ConversationMember);
        }

        const chatData: CreateChatPayload = {
          chatType: uniqueEmails.length === 1 ? "oneOnOne" : "group",
          members,
        };

        if (topic && uniqueEmails.length > 1) {
          chatData.topic = topic;
        }

        const newChat = (await client.api("/chats").post(chatData)) as Chat;

        return {
          content: [
            {
              type: "text",
              text: `✅ Chat created successfully. Chat ID: ${newChat?.id}`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Rename a group chat
  server.tool(
    "rename_chat",
    "Rename a group chat by changing its topic (title). Only group chats can be renamed — 1:1 (oneOnOne) chats have no topic. The topic is limited to 250 characters and cannot contain ':'.",
    {
      chatId: z.string().describe("Chat ID of the group chat to rename"),
      topic: z
        .string()
        .min(1)
        .max(250)
        .describe("New chat topic (title), up to 250 characters; ':' is not allowed"),
    },
    async ({ chatId, topic }) => {
      try {
        if (topic.includes(":")) {
          return {
            content: [
              {
                type: "text",
                text: "❌ Error: Chat topic cannot contain ':' (not allowed by Microsoft Graph).",
              },
            ],
          };
        }

        const client = await graphService.getClient();

        await client.api(`/chats/${chatId}`).patch({ topic });

        return {
          content: [
            {
              type: "text",
              text: `✅ Chat renamed successfully. New topic: "${topic}"`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Failed to rename chat: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Update/Edit a chat message
  server.tool(
    "update_chat_message",
    "Update (edit) a chat message that was previously sent. Only the message sender can update their own messages. Supports updating content with text or Markdown formatting, mentions, and importance levels.",
    {
      chatId: z.string().describe("Chat ID"),
      messageId: z.string().describe("Message ID to update"),
      message: z.string().describe("New message content"),
      importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
      format: z.enum(["text", "markdown"]).optional().describe("Message format (text or markdown)"),
      mentions: z
        .array(
          z.object({
            mention: z
              .string()
              .describe("The @mention text (e.g., 'john.doe' or 'john.doe@company.com')"),
            userId: z.string().describe("Azure AD User ID of the mentioned user"),
          })
        )
        .optional()
        .describe("Array of @mentions to include in the message"),
    },
    async ({ chatId, messageId, message, importance, format = "text", mentions }) => {
      try {
        const client = await graphService.getClient();

        // Process message content based on format
        let content: string;
        let contentType: "text" | "html";

        if (format === "markdown") {
          content = await markdownToHtml(message);
          contentType = "html";
        } else {
          content = message;
          contentType = "text";
        }

        // Process @mentions if provided
        const mentionMappings: Array<{ mention: string; userId: string; displayName: string }> = [];
        if (mentions && mentions.length > 0) {
          // Convert provided mentions to mappings with display names
          for (const mention of mentions) {
            try {
              // Get user info to get display name
              const userResponse = await client
                .api(`/users/${mention.userId}`)
                .select("displayName")
                .get();
              mentionMappings.push({
                mention: mention.mention,
                userId: mention.userId,
                displayName: userResponse.displayName || mention.mention,
              });
            } catch (_error) {
              console.warn(
                `Could not resolve user ${mention.userId}, using mention text as display name`
              );
              mentionMappings.push({
                mention: mention.mention,
                userId: mention.userId,
                displayName: mention.mention,
              });
            }
          }
        }

        // Process mentions in HTML content
        let finalMentions: Array<{
          id: number;
          mentionText: string;
          mentioned: { user: { id: string } };
        }> = [];
        if (mentionMappings.length > 0) {
          const result = processMentionsInHtml(content, mentionMappings);
          content = result.content;
          finalMentions = result.mentions;

          // Ensure we're using HTML content type when mentions are present
          contentType = "html";
        }

        // Build message payload for update
        const messagePayload: any = {
          body: {
            content,
            contentType,
          },
        };

        if (importance) {
          messagePayload.importance = importance;
        }

        if (finalMentions.length > 0) {
          messagePayload.mentions = finalMentions;
        }

        // Update the message using PATCH
        // Note: Using /me/chats/ endpoint for delegated permissions
        // The API also requires proper permissions: Chat.ReadWrite
        await client.api(`/me/chats/${chatId}/messages/${messageId}`).patch(messagePayload);

        // Build success message
        const successText = `✅ Message updated successfully. Message ID: ${messageId}${
          finalMentions.length > 0
            ? `\n📱 Mentions: ${finalMentions.map((m) => m.mentionText).join(", ")}`
            : ""
        }`;

        return {
          content: [
            {
              type: "text" as const,
              text: successText,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to update message: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Soft delete a chat message
  server.tool(
    "delete_chat_message",
    "Soft delete a chat message that was previously sent. Only the message sender can delete their own messages. The message will be marked as deleted but can still be seen as '[This message has been deleted]'.",
    {
      chatId: z.string().describe("Chat ID"),
      messageId: z.string().describe("Message ID to delete"),
    },
    async ({ chatId, messageId }) => {
      try {
        const client = await graphService.getClient();

        // Get current user ID for the endpoint
        const me = (await client.api("/me").get()) as { id: string };

        // Soft delete the message using POST
        // Endpoint: POST /users/{userId}/chats/{chatsId}/messages/{chatMessageId}/softDelete
        await client
          .api(`/users/${me.id}/chats/${chatId}/messages/${messageId}/softDelete`)
          .post({});

        return {
          content: [
            {
              type: "text" as const,
              text: `✅ Message deleted successfully. Message ID: ${messageId}`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to delete message: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Set a reaction on a chat message
  server.tool(
    "set_chat_message_reaction",
    "Add a reaction to a message in a chat conversation. Supports Unicode emoji characters and named reactions (like, angry, sad, laugh, heart, surprised).",
    {
      chatId: z.string().describe("Chat ID"),
      messageId: z.string().describe("Message ID to react to"),
      reactionType: z
        .string()
        .describe(
          'Reaction type - Unicode emoji (e.g., "👍") or named reaction (e.g., "like", "heart")'
        ),
    },
    async ({ chatId, messageId, reactionType }) => {
      try {
        const client = await graphService.getClient();

        await client
          .api(`/chats/${chatId}/messages/${messageId}/setReaction`)
          .post({ reactionType });

        return {
          content: [
            {
              type: "text" as const,
              text: `✅ Reaction ${reactionType} added to message ${messageId}.`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to set reaction: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Unset a reaction on a chat message
  server.tool(
    "unset_chat_message_reaction",
    "Remove a reaction from a message in a chat conversation.",
    {
      chatId: z.string().describe("Chat ID"),
      messageId: z.string().describe("Message ID to remove reaction from"),
      reactionType: z
        .string()
        .describe(
          'Reaction type to remove - Unicode emoji (e.g., "👍") or named reaction (e.g., "like", "heart")'
        ),
    },
    async ({ chatId, messageId, reactionType }) => {
      try {
        const client = await graphService.getClient();

        await client
          .api(`/chats/${chatId}/messages/${messageId}/unsetReaction`)
          .post({ reactionType });

        return {
          content: [
            {
              type: "text" as const,
              text: `✅ Reaction ${reactionType} removed from message ${messageId}.`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to unset reaction: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Send a file to a chat
  server.tool(
    "send_file_to_chat",
    "Send a file as a message to a Teams chat. Provide either filePath (a file on the MCP server's filesystem) or driveItemId (a file already uploaded to OneDrive via create_file_upload_session — use this to send files from the caller's machine). The file is sent as a reference attachment.",
    {
      chatId: z.string().describe("Chat ID"),
      filePath: z
        .string()
        .optional()
        .describe("Path to a file on the MCP server's filesystem to upload"),
      driveItemId: z
        .string()
        .optional()
        .describe(
          "ID of a drive item already uploaded via create_file_upload_session (alternative to filePath)"
        ),
      message: z.string().optional().describe("Optional message text to accompany the file"),
      fileName: z
        .string()
        .optional()
        .describe("Optional custom filename (defaults to the original file name)"),
      format: z.enum(["text", "markdown"]).optional().describe("Message format (text or markdown)"),
      importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
    },
    async ({
      chatId,
      filePath,
      driveItemId,
      message,
      fileName,
      format = "text",
      importance = "normal",
    }) => {
      try {
        if (filePath && driveItemId) {
          return {
            content: [
              {
                type: "text" as const,
                text: "❌ Provide either filePath or driveItemId, not both.",
              },
            ],
            isError: true,
          };
        }

        const client = await graphService.getClient();

        let uploadResult: FileUploadResult;
        if (driveItemId) {
          uploadResult = await resolveChatDriveItem(graphService, driveItemId);
        } else if (filePath) {
          uploadResult = await uploadFileToChat(graphService, filePath, fileName);
        } else {
          return {
            content: [
              {
                type: "text" as const,
                text: "❌ Provide filePath (file on the server) or driveItemId (file uploaded via create_file_upload_session).",
              },
            ],
            isError: true,
          };
        }

        // Build message content — must be HTML with attachment reference tag
        let content = "";
        if (message) {
          if (format === "markdown") {
            content = await markdownToHtml(message);
          } else {
            content = escapeHtml(message);
          }
        }

        const attachmentTag = `<attachment id="${uploadResult.attachmentId}"></attachment>`;
        content = content ? `${content}<br>${attachmentTag}` : attachmentTag;

        const attachments = buildFileAttachment(uploadResult);
        const messagePayload: any = {
          body: { content, contentType: "html" },
          importance,
          attachments,
        };

        const result = (await client
          .api(`/me/chats/${chatId}/messages`)
          .post(messagePayload)) as ChatMessage;

        return {
          content: [
            {
              type: "text" as const,
              text: `✅ File sent successfully to chat.\nFile: ${uploadResult.fileName} (${formatFileSize(uploadResult.fileSize)})\nMessage ID: ${result.id}`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to send file: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Add a member to a group chat
  server.tool(
    "add_chat_member",
    "Add a user to a group chat. Optionally controls how much of the existing chat history the new member can see (default: all of it).",
    {
      chatId: z.string().describe("Chat ID of the group chat"),
      userEmail: z.string().describe("Email address of the user to add"),
      shareHistory: z
        .enum(["all", "none"])
        .optional()
        .describe("How much chat history the new member sees (default: all)"),
      shareHistorySince: z
        .string()
        .optional()
        .describe("Share history starting from this ISO datetime (overrides shareHistory)"),
    },
    async ({ chatId, userEmail, shareHistory = "all", shareHistorySince }) => {
      try {
        const client = await graphService.getClient();

        const user = (await client.api(`/users/${userEmail}`).get()) as User;

        // This endpoint requires the user@odata.bind reference format.
        // Omitting visibleHistoryStartDateTime means "share no history";
        // 0001-01-01T00:00:00Z is Graph's marker for "share everything".
        const memberPayload: Record<string, unknown> = {
          "@odata.type": "#microsoft.graph.aadUserConversationMember",
          roles: ["owner"],
          "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${user?.id}')`,
        };
        if (shareHistorySince) {
          memberPayload.visibleHistoryStartDateTime = shareHistorySince;
        } else if (shareHistory === "all") {
          memberPayload.visibleHistoryStartDateTime = "0001-01-01T00:00:00Z";
        }

        await client.api(`/chats/${chatId}/members`).post(memberPayload);

        const historyNote = shareHistorySince
          ? `history visible since ${shareHistorySince}`
          : shareHistory === "all"
            ? "full history visible"
            : "no history visible";
        return {
          content: [
            {
              type: "text" as const,
              text: `✅ ${userEmail} added to chat (${historyNote}).`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to add chat member: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Remove a member from a group chat
  server.tool(
    "remove_chat_member",
    "Remove a user from a group chat.",
    {
      chatId: z.string().describe("Chat ID of the group chat"),
      userEmail: z.string().describe("Email address of the member to remove"),
    },
    async ({ chatId, userEmail }) => {
      try {
        const client = await graphService.getClient();

        // Resolve the user and match members by user id — a member's email
        // can differ from the address the caller knows.
        const user = (await client.api(`/users/${userEmail}`).get()) as User;

        const response = (await client
          .api(`/chats/${chatId}/members`)
          .get()) as GraphApiResponse<AadUserConversationMember>;
        const target = response?.value?.find((member) => member.userId === user?.id);

        if (!target?.id) {
          return {
            content: [
              {
                type: "text" as const,
                text: `❌ ${userEmail} is not a member of this chat.`,
              },
            ],
            isError: true,
          };
        }

        await client.api(`/chats/${chatId}/members/${target.id}`).delete();

        return {
          content: [
            {
              type: "text" as const,
              text: `✅ ${userEmail} removed from chat.`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to remove chat member: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Leave a group chat
  server.tool(
    "leave_chat",
    "Leave a group chat (removes the current user from the chat).",
    {
      chatId: z.string().describe("Chat ID of the group chat to leave"),
    },
    async ({ chatId }) => {
      try {
        const client = await graphService.getClient();

        const me = (await client.api("/me").get()) as User;
        const response = (await client
          .api(`/chats/${chatId}/members`)
          .get()) as GraphApiResponse<AadUserConversationMember>;
        const own = response?.value?.find((member) => member.userId === me?.id);

        if (!own?.id) {
          return {
            content: [
              {
                type: "text" as const,
                text: "❌ You are not a member of this chat.",
              },
            ],
            isError: true,
          };
        }

        await client.api(`/chats/${chatId}/members/${own.id}`).delete();

        return {
          content: [
            {
              type: "text" as const,
              text: "✅ You left the chat.",
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to leave chat: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Pin a message in a chat
  server.tool(
    "pin_chat_message",
    "Pin a message in a chat so it is shown at the top of the conversation for all members.",
    {
      chatId: z.string().describe("Chat ID"),
      messageId: z.string().describe("Message ID to pin"),
    },
    async ({ chatId, messageId }) => {
      try {
        const client = await graphService.getClient();

        await client.api(`/chats/${chatId}/pinnedMessages`).post({
          "message@odata.bind": `https://graph.microsoft.com/v1.0/chats/${chatId}/messages/${messageId}`,
        });

        return {
          content: [
            {
              type: "text" as const,
              text: `✅ Message ${messageId} pinned.`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to pin message: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Unpin a message in a chat
  server.tool(
    "unpin_chat_message",
    "Unpin a previously pinned message in a chat.",
    {
      chatId: z.string().describe("Chat ID"),
      messageId: z.string().describe("Message ID to unpin"),
    },
    async ({ chatId, messageId }) => {
      try {
        const client = await graphService.getClient();

        // Graph uses the message id as the pinned-item id.
        await client.api(`/chats/${chatId}/pinnedMessages/${messageId}`).delete();

        return {
          content: [
            {
              type: "text" as const,
              text: `✅ Message ${messageId} unpinned.`,
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to unpin message: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Hide a chat from the current user's chat list
  server.tool(
    "hide_chat",
    "Hide a chat from your own chat list (does not leave or delete the chat, and does not affect other members). The chat reappears automatically when there is new activity.",
    {
      chatId: z.string().describe("Chat ID to hide"),
    },
    async ({ chatId }) => {
      try {
        const client = await graphService.getClient();

        // hideForUser requires the user's id and tenantId; take them from our
        // own membership entry in the chat.
        const me = (await client.api("/me").get()) as User;
        const response = (await client
          .api(`/chats/${chatId}/members`)
          .get()) as GraphApiResponse<AadUserConversationMember>;
        const own = response?.value?.find((member) => member.userId === me?.id);

        if (!own) {
          return {
            content: [
              {
                type: "text" as const,
                text: "❌ You are not a member of this chat.",
              },
            ],
            isError: true,
          };
        }

        await client.api(`/chats/${chatId}/hideForUser`).post({
          user: { id: me?.id, tenantId: own.tenantId },
        });

        return {
          content: [
            {
              type: "text" as const,
              text: "✅ Chat hidden from your chat list. It reappears on new activity.",
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to hide chat: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Unhide a chat in the current user's chat list
  server.tool(
    "unhide_chat",
    "Unhide a previously hidden chat in your own chat list.",
    {
      chatId: z.string().describe("Chat ID to unhide"),
    },
    async ({ chatId }) => {
      try {
        const client = await graphService.getClient();

        const me = (await client.api("/me").get()) as User;
        const response = (await client
          .api(`/chats/${chatId}/members`)
          .get()) as GraphApiResponse<AadUserConversationMember>;
        const own = response?.value?.find((member) => member.userId === me?.id);

        if (!own) {
          return {
            content: [
              {
                type: "text" as const,
                text: "❌ You are not a member of this chat.",
              },
            ],
            isError: true,
          };
        }

        await client.api(`/chats/${chatId}/unhideForUser`).post({
          user: { id: me?.id, tenantId: own.tenantId },
        });

        return {
          content: [
            {
              type: "text" as const,
              text: "✅ Chat is visible in your chat list again.",
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to unhide chat: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Create a resumable upload session so a client can upload a file directly
  server.tool(
    "create_file_upload_session",
    "Create a resumable OneDrive/SharePoint upload session for sending a file to Teams from the caller's machine. Returns a pre-authenticated uploadUrl: PUT the file bytes to it directly (no Authorization header needed) — the MCP server never stores the file. After the final chunk Graph returns the driveItem JSON; pass its 'id' as driveItemId to send_file_to_chat or send_file_to_channel.",
    {
      fileName: z.string().describe("Name for the uploaded file"),
      target: z
        .enum(["chat", "channel"])
        .describe(
          "Where the file will be sent afterwards: 'chat' uploads to your OneDrive, 'channel' to the channel's SharePoint folder"
        ),
      teamId: z.string().optional().describe("Team ID (required when target is 'channel')"),
      channelId: z.string().optional().describe("Channel ID (required when target is 'channel')"),
    },
    async ({ fileName, target, teamId, channelId }) => {
      try {
        let session: UploadSessionInfo;
        if (target === "channel") {
          if (!teamId || !channelId) {
            return {
              content: [
                {
                  type: "text" as const,
                  text: "❌ teamId and channelId are required when target is 'channel'.",
                },
              ],
              isError: true,
            };
          }
          session = await createUploadSessionForChannel(graphService, teamId, channelId, fileName);
        } else {
          session = await createUploadSessionForChat(graphService, fileName);
        }

        const guide = {
          uploadUrl: session.uploadUrl,
          expiresAt: session.expirationDateTime,
          target,
          fileName,
          rules: {
            authorization: "none — the uploadUrl is pre-authenticated",
            singlePutMaxBytes: 62914560,
            chunkSizeMultipleOfBytes: 327680,
            contentRange: "required, e.g. 'bytes 0-{size-1}/{size}' for a single PUT",
            finalResponse:
              "JSON driveItem — pass its 'id' to send_file_to_chat/send_file_to_channel as driveItemId",
          },
          examples: {
            bash: 'SIZE=$(stat -c%s "$FILE" 2>/dev/null || stat -f%z "$FILE"); curl -sS -X PUT "$UPLOAD_URL" -H "Content-Range: bytes 0-$((SIZE-1))/$SIZE" --data-binary @"$FILE"',
            powershell:
              '$size=(Get-Item $file).Length; curl.exe -sS -X PUT $uploadUrl -H "Content-Range: bytes 0-$($size-1)/$size" --data-binary "@$file"',
            largeFiles:
              "over 60 MiB: upload sequential chunks (multiples of 320 KiB) with matching Content-Range headers",
          },
        };

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(guide, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to create upload session: ${errorMessage}`,
            },
          ],
          isError: true,
        };
      }
    }
  );
}
