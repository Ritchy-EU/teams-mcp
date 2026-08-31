/**
 * Upload image as hosted content for Teams messages
 * This creates a temporary hosted content that can be referenced in message attachments
 */
export async function uploadImageAsHostedContent(graphService, teamId, channelId, imageData, contentType, fileName) {
    try {
        const client = await graphService.getClient();
        // Convert Buffer to base64 if needed
        const contentBytes = typeof imageData === "string" ? imageData : imageData.toString("base64");
        // Create hosted content
        const hostedContent = {
            "@microsoft.graph.temporaryId": `temp_${Date.now()}_${Math.random().toString(36).substring(7)}`,
            contentBytes,
            contentType,
        };
        // Upload hosted content
        const response = await client
            .api(`/teams/${teamId}/channels/${channelId}/messages/hostedContents`)
            .post(hostedContent);
        const hostedContentId = response.id;
        // Create attachment reference
        const attachment = {
            id: hostedContentId,
            contentType,
            contentUrl: `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages/hostedContents/${hostedContentId}/$value`,
            name: fileName || `image.${getFileExtensionFromMimeType(contentType)}`,
        };
        return { hostedContentId, attachment };
    }
    catch (error) {
        console.error("Error uploading image as hosted content:", error);
        return null;
    }
}
/**
 * Validate image content type
 */
export function isValidImageType(contentType) {
    const validTypes = [
        "image/jpeg",
        "image/jpg",
        "image/png",
        "image/gif",
        "image/webp",
        "image/bmp",
        "image/svg+xml",
    ];
    return validTypes.includes(contentType.toLowerCase());
}
/**
 * Get file extension from MIME type
 */
export function getFileExtensionFromMimeType(mimeType) {
    const extensions = {
        "image/jpeg": "jpg",
        "image/jpg": "jpg",
        "image/png": "png",
        "image/gif": "gif",
        "image/webp": "webp",
        "image/bmp": "bmp",
        "image/svg+xml": "svg",
    };
    return extensions[mimeType.toLowerCase()] || "img";
}
/**
 * Convert image URL to base64 for upload
 */
export async function imageUrlToBase64(imageUrl) {
    try {
        const response = await fetch(imageUrl);
        if (!response.ok) {
            throw new Error(`Failed to fetch image: ${response.statusText}`);
        }
        const contentType = response.headers.get("content-type") || "image/jpeg";
        if (!isValidImageType(contentType)) {
            throw new Error(`Unsupported image type: ${contentType}`);
        }
        const arrayBuffer = await response.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);
        const base64Data = buffer.toString("base64");
        return {
            data: base64Data,
            contentType,
        };
    }
    catch (error) {
        console.error("Error converting image URL to base64:", error);
        return null;
    }
}
/**
 * Extracts a minimal attachment summary from Graph API ChatMessageAttachment array.
 * Returns undefined if there are no meaningful attachments to report.
 */
export function extractAttachmentSummaries(attachments) {
    if (!attachments || attachments.length === 0) {
        return undefined;
    }
    const summaries = [];
    for (const att of attachments) {
        if (!att.id && !att.name && !att.contentUrl)
            continue;
        summaries.push({
            id: att.id ?? undefined,
            name: att.name ?? undefined,
            contentType: att.contentType ?? undefined,
            contentUrl: att.contentUrl ?? undefined,
            thumbnailUrl: att.thumbnailUrl ?? undefined,
        });
    }
    return summaries.length > 0 ? summaries : undefined;
}
/** Extract hosted content ids (inline images) referenced in a message's HTML body. */
export function extractHostedContentIds(html) {
    const regex = /hostedContents\/([a-zA-Z0-9_=-]+)\/\$value/g;
    const ids = new Set();
    let match = regex.exec(html);
    while (match !== null) {
        ids.add(match[1]);
        match = regex.exec(html);
    }
    return [...ids];
}
/**
 * Collect all downloadable content of a message as attachment summaries:
 * regular file/card attachments plus inline images (hosted content), the
 * latter marked with contentType "hostedContent" so callers know to fetch
 * them as bytes rather than by URL.
 */
export function collectMessageAttachments(attachments, bodyHtml) {
    const files = extractAttachmentSummaries(attachments) ?? [];
    const inline = bodyHtml
        ? extractHostedContentIds(bodyHtml).map((id) => ({ id, contentType: "hostedContent" }))
        : [];
    const all = [...files, ...inline];
    return all.length > 0 ? all : undefined;
}
//# sourceMappingURL=attachments.js.map