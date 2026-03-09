import { z } from "zod";
import { markdownToHtml } from "../utils/markdown.js";
import { processMentionsInHtml } from "../utils/users.js";
/**
 * Registers all chat-related MCP tools on the given server.
 * Tools include: list_chats, get_chat_messages, send_chat_message,
 * create_chat, update_chat_message, and delete_chat_message.
 *
 * @param server - The MCP server instance to register tools on.
 * @param graphService - The Microsoft Graph service used for API calls.
 */
export function registerChatTools(server, graphService) {
    // List user's chats
    server.tool("list_chats", "List all recent chats (1:1 conversations and group chats) that the current user participates in. Returns chat topics, types, and participant information.", {}, async () => {
        try {
            // Build query parameters
            const queryParams = ["$expand=members"];
            const queryString = queryParams.join("&");
            const client = await graphService.getClient();
            const response = (await client
                .api(`/me/chats?${queryString}`)
                .get());
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
            const chatList = response.value.map((chat) => ({
                id: chat.id,
                topic: chat.topic || "No topic",
                chatType: chat.chatType,
                members: chat.members?.map((member) => member.displayName).join(", ") ||
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
        }
        catch (error) {
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
    });
    // Get chat messages with pagination support
    server.tool("get_chat_messages", "Retrieve recent messages from a specific chat conversation. Returns message content, sender information, and timestamps.", {
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
            .describe("Fetch all messages using pagination (up to limit). When true, follows @odata.nextLink to get more messages."),
    }, async ({ chatId, limit, since, until, fromUser, orderBy, descending, fetchAll }) => {
        try {
            const client = await graphService.getClient();
            // Apply defaults for parameters (in case Zod validation is bypassed)
            const effectiveLimit = limit ?? 20;
            const effectiveOrderBy = orderBy ?? "createdDateTime";
            const effectiveDescending = descending ?? true;
            const effectiveFetchAll = fetchAll ?? false;
            // Build query parameters - use smaller page size for pagination
            const pageSize = effectiveFetchAll ? 50 : Math.min(effectiveLimit, 50);
            const queryParams = [`$top=${pageSize}`];
            // Add ordering - Graph API only supports descending order for datetime fields in chat messages
            if ((effectiveOrderBy === "createdDateTime" || effectiveOrderBy === "lastModifiedDateTime") &&
                !effectiveDescending) {
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
            const filters = [];
            if (fromUser) {
                filters.push(`from/user/id eq '${fromUser}'`);
            }
            if (filters.length > 0) {
                queryParams.push(`$filter=${filters.join(" and ")}`);
            }
            const queryString = queryParams.join("&");
            // Fetch messages with pagination support
            const allMessages = [];
            let nextLink;
            let pageCount = 0;
            const maxPages = 100; // Safety limit to prevent infinite loops
            // First request
            let response = (await client
                .api(`/me/chats/${chatId}/messages?${queryString}`)
                .get());
            if (response?.value) {
                allMessages.push(...response.value);
            }
            // Follow pagination if fetchAll is enabled
            if (effectiveFetchAll) {
                nextLink = response["@odata.nextLink"];
                while (nextLink && allMessages.length < effectiveLimit && pageCount < maxPages) {
                    pageCount++;
                    try {
                        response = (await client.api(nextLink).get());
                        if (response?.value) {
                            allMessages.push(...response.value);
                        }
                        nextLink = response["@odata.nextLink"];
                    }
                    catch (pageError) {
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
                filteredMessages = allMessages.filter((message) => {
                    if (!message.createdDateTime)
                        return true;
                    const messageDate = new Date(message.createdDateTime);
                    if (since) {
                        const sinceDate = new Date(since);
                        if (messageDate <= sinceDate)
                            return false;
                    }
                    if (until) {
                        const untilDate = new Date(until);
                        if (messageDate >= untilDate)
                            return false;
                    }
                    return true;
                });
            }
            // Apply limit after filtering
            const limitedMessages = filteredMessages.slice(0, effectiveLimit);
            const messageList = limitedMessages.map((message) => {
                const summary = {
                    id: message.id,
                    content: message.body?.content,
                    from: message.from?.user?.displayName,
                    createdDateTime: message.createdDateTime,
                };
                // Include attachment metadata if present
                if (message.attachments && message.attachments.length > 0) {
                    summary.attachments = message.attachments.map((att) => ({
                        id: att.id,
                        name: att.name,
                        contentType: att.contentType,
                        contentUrl: att.contentUrl,
                    }));
                }
                return summary;
            });
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            filters: { since, until, fromUser },
                            filteringMethod: since || until ? "client-side" : "server-side",
                            paginationEnabled: fetchAll,
                            pagesRetrieved: pageCount + 1,
                            totalRetrieved: allMessages.length,
                            totalReturned: messageList.length,
                            hasMore: !!response["@odata.nextLink"] || filteredMessages.length > limit,
                            messages: messageList,
                        }, null, 2),
                    },
                ],
            };
        }
        catch (error) {
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
    });
    // Send chat message
    server.tool("send_chat_message", "Send a message to a specific chat conversation. Supports text and markdown formatting, mentions, and importance levels.", {
        chatId: z.string().describe("Chat ID"),
        message: z.string().describe("Message content"),
        importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
        format: z.enum(["text", "markdown"]).optional().describe("Message format (text or markdown)"),
        mentions: z
            .array(z.object({
            mention: z
                .string()
                .describe("The @mention text (e.g., 'john.doe' or 'john.doe@company.com')"),
            userId: z.string().describe("Azure AD User ID of the mentioned user"),
        }))
            .optional()
            .describe("Array of @mentions to include in the message"),
    }, async ({ chatId, message, importance = "normal", format = "text", mentions }) => {
        try {
            const client = await graphService.getClient();
            // Process message content based on format
            let content;
            let contentType;
            if (format === "markdown") {
                content = await markdownToHtml(message);
                contentType = "html";
            }
            else {
                content = message;
                contentType = "text";
            }
            // Process @mentions if provided
            const mentionMappings = [];
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
                    }
                    catch (_error) {
                        console.warn(`Could not resolve user ${mention.userId}, using mention text as display name`);
                        mentionMappings.push({
                            mention: mention.mention,
                            userId: mention.userId,
                            displayName: mention.mention,
                        });
                    }
                }
            }
            // Process mentions in HTML content
            let finalMentions = [];
            if (mentionMappings.length > 0) {
                const result = processMentionsInHtml(content, mentionMappings);
                content = result.content;
                finalMentions = result.mentions;
                // Ensure we're using HTML content type when mentions are present
                contentType = "html";
            }
            // Build message payload
            const messagePayload = {
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
                .post(messagePayload));
            // Build success message
            const successText = `✅ Message sent successfully. Message ID: ${result.id}${finalMentions.length > 0
                ? `\n📱 Mentions: ${finalMentions.map((m) => m.mentionText).join(", ")}`
                : ""}`;
            return {
                content: [
                    {
                        type: "text",
                        text: successText,
                    },
                ],
            };
        }
        catch (error) {
            return {
                content: [
                    {
                        type: "text",
                        text: `❌ Failed to send message: ${error.message}`,
                    },
                ],
                isError: true,
            };
        }
    });
    // Create new chat (1:1 or group)
    server.tool("create_chat", "Create a new chat conversation. Can be a 1:1 chat (with one other user) or a group chat (with multiple users). Group chats can optionally have a topic.", {
        userEmails: z.array(z.string()).describe("Array of user email addresses to add to chat"),
        topic: z.string().optional().describe("Chat topic (for group chats)"),
    }, async ({ userEmails, topic }) => {
        try {
            const client = await graphService.getClient();
            // Get current user ID
            const me = (await client.api("/me").get());
            // Create members array
            const members = [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    user: {
                        id: me?.id,
                    },
                    roles: ["owner"],
                },
            ];
            // Add other users as members
            for (const email of userEmails) {
                const user = (await client.api(`/users/${email}`).get());
                members.push({
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    user: {
                        id: user?.id,
                    },
                    roles: ["member"],
                });
            }
            const chatData = {
                chatType: userEmails.length === 1 ? "oneOnOne" : "group",
                members,
            };
            if (topic && userEmails.length > 1) {
                chatData.topic = topic;
            }
            const newChat = (await client.api("/chats").post(chatData));
            return {
                content: [
                    {
                        type: "text",
                        text: `✅ Chat created successfully. Chat ID: ${newChat?.id}`,
                    },
                ],
            };
        }
        catch (error) {
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
    });
    // Update/Edit a chat message
    server.tool("update_chat_message", "Update (edit) a chat message that was previously sent. Only the message sender can update their own messages. Supports updating content with text or Markdown formatting, mentions, and importance levels.", {
        chatId: z.string().describe("Chat ID"),
        messageId: z.string().describe("Message ID to update"),
        message: z.string().describe("New message content"),
        importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
        format: z.enum(["text", "markdown"]).optional().describe("Message format (text or markdown)"),
        mentions: z
            .array(z.object({
            mention: z
                .string()
                .describe("The @mention text (e.g., 'john.doe' or 'john.doe@company.com')"),
            userId: z.string().describe("Azure AD User ID of the mentioned user"),
        }))
            .optional()
            .describe("Array of @mentions to include in the message"),
    }, async ({ chatId, messageId, message, importance, format = "text", mentions }) => {
        try {
            const client = await graphService.getClient();
            // Process message content based on format
            let content;
            let contentType;
            if (format === "markdown") {
                content = await markdownToHtml(message);
                contentType = "html";
            }
            else {
                content = message;
                contentType = "text";
            }
            // Process @mentions if provided
            const mentionMappings = [];
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
                    }
                    catch (_error) {
                        console.warn(`Could not resolve user ${mention.userId}, using mention text as display name`);
                        mentionMappings.push({
                            mention: mention.mention,
                            userId: mention.userId,
                            displayName: mention.mention,
                        });
                    }
                }
            }
            // Process mentions in HTML content
            let finalMentions = [];
            if (mentionMappings.length > 0) {
                const result = processMentionsInHtml(content, mentionMappings);
                content = result.content;
                finalMentions = result.mentions;
                // Ensure we're using HTML content type when mentions are present
                contentType = "html";
            }
            // Build message payload for update
            const messagePayload = {
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
            const successText = `✅ Message updated successfully. Message ID: ${messageId}${finalMentions.length > 0
                ? `\n📱 Mentions: ${finalMentions.map((m) => m.mentionText).join(", ")}`
                : ""}`;
            return {
                content: [
                    {
                        type: "text",
                        text: successText,
                    },
                ],
            };
        }
        catch (error) {
            const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
            return {
                content: [
                    {
                        type: "text",
                        text: `❌ Failed to update message: ${errorMessage}`,
                    },
                ],
                isError: true,
            };
        }
    });
    // Soft delete a chat message
    server.tool("delete_chat_message", "Soft delete a chat message that was previously sent. Only the message sender can delete their own messages. The message will be marked as deleted but can still be seen as '[This message has been deleted]'.", {
        chatId: z.string().describe("Chat ID"),
        messageId: z.string().describe("Message ID to delete"),
    }, async ({ chatId, messageId }) => {
        try {
            const client = await graphService.getClient();
            // Get current user ID for the endpoint
            const me = (await client.api("/me").get());
            // Soft delete the message using POST
            // Endpoint: POST /users/{userId}/chats/{chatsId}/messages/{chatMessageId}/softDelete
            await client
                .api(`/users/${me.id}/chats/${chatId}/messages/${messageId}/softDelete`)
                .post({});
            return {
                content: [
                    {
                        type: "text",
                        text: `✅ Message deleted successfully. Message ID: ${messageId}`,
                    },
                ],
            };
        }
        catch (error) {
            const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
            return {
                content: [
                    {
                        type: "text",
                        text: `❌ Failed to delete message: ${errorMessage}`,
                    },
                ],
                isError: true,
            };
        }
    });
    // Download attachments and hosted content from chat messages
    server.tool("download_chat_attachment", "Download files and images from a chat message. Handles both inline images (hosted content) and file attachments (OneDrive/SharePoint references). Use get_chat_messages first to see available attachments.", {
        chatId: z.string().describe("Chat ID"),
        messageId: z.string().describe("Message ID containing the attachment"),
        attachmentIndex: z
            .number()
            .min(0)
            .optional()
            .describe("Index of a specific attachment to download (0-based). If not provided, downloads all attachments."),
        savePath: z
            .string()
            .optional()
            .describe("Optional file path to save the content. Supports UNC paths (e.g., \\\\wsl.localhost\\Ubuntu\\tmp\\file.png)."),
    }, async ({ chatId, messageId, attachmentIndex, savePath }) => {
        try {
            const client = await graphService.getClient();
            // Fetch the message to inspect attachments and body
            const message = (await client
                .api(`/me/chats/${chatId}/messages/${messageId}`)
                .get());
            if (!message) {
                return {
                    content: [{ type: "text", text: "❌ Message not found." }],
                    isError: true,
                };
            }
            const items = [];
            // 1. Extract hosted content IDs from message body HTML (inline images)
            const bodyContent = message.body?.content || "";
            const hostedContentRegex = /hostedContents\/([a-zA-Z0-9_=-]+)\/\$value|itemid="([^"]+)"/gi;
            let match;
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
            const targetItems = attachmentIndex !== undefined ? [items[attachmentIndex]].filter(Boolean) : items;
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
            const results = [];
            for (const item of targetItems) {
                const itemIndex = items.indexOf(item);
                try {
                    let buffer;
                    if (item.type === "hostedContent") {
                        // Download hosted content (inline images)
                        const response = await client
                            .api(`/chats/${chatId}/messages/${messageId}/hostedContents/${item.id}/$value`)
                            .responseType("arraybuffer")
                            .get();
                        buffer = Buffer.from(response);
                    }
                    else {
                        // Download file reference via Shares API
                        // Encode the SharePoint URL as base64url with "u!" prefix
                        const encodedUrl = `u!${Buffer.from(item.contentUrl)
                            .toString("base64")
                            .replace(/\+/g, "-")
                            .replace(/\//g, "_")
                            .replace(/=+$/, "")}`;
                        const response = await client
                            .api(`/shares/${encodedUrl}/driveItem/content`)
                            .responseType("arraybuffer")
                            .get();
                        buffer = Buffer.from(response);
                    }
                    const result = {
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
                        const isUncPath = normalizedPath.startsWith("\\\\") || normalizedPath.startsWith("//");
                        let finalPath = normalizedPath;
                        if (targetItems.length > 1) {
                            const ext = path.extname(normalizedPath);
                            const base = ext ? normalizedPath.slice(0, -ext.length) : normalizedPath;
                            finalPath = `${base}_${itemIndex}${ext || path.extname(item.name)}`;
                        }
                        const targetPath = isUncPath ? finalPath : path.resolve(finalPath);
                        await fs.writeFile(targetPath, buffer);
                        result.savedTo = targetPath;
                    }
                    else {
                        result.base64Data = buffer.toString("base64");
                    }
                    results.push(result);
                }
                catch (downloadError) {
                    const errorMsg = downloadError instanceof Error ? downloadError.message : "Unknown error";
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
                        text: JSON.stringify({
                            summary,
                            messageId,
                            totalAttachments: items.length,
                            downloaded: targetItems.length,
                            successCount,
                            errorCount,
                            attachments: results,
                        }, null, 2),
                    },
                ],
            };
        }
        catch (error) {
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
    });
}
//# sourceMappingURL=chats.js.map