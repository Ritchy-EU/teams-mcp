import { promises as fs } from "node:fs";
import { basename, extname } from "node:path";
/** Simple upload threshold (4 MB) */
const SIMPLE_UPLOAD_MAX_SIZE = 4 * 1024 * 1024;
/** Upload session chunk size — must be a multiple of 320 KiB */
const UPLOAD_CHUNK_SIZE = 320 * 1024 * 10; // 3.2 MB
/** Extension → MIME type map for common file types */
const MIME_TYPES = {
    ".pdf": "application/pdf",
    ".doc": "application/msword",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xls": "application/vnd.ms-excel",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".ppt": "application/vnd.ms-powerpoint",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".zip": "application/zip",
    ".7z": "application/x-7z-compressed",
    ".rar": "application/vnd.rar",
    ".tar": "application/x-tar",
    ".gz": "application/gzip",
    ".txt": "text/plain",
    ".csv": "text/csv",
    ".json": "application/json",
    ".xml": "application/xml",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".svg": "image/svg+xml",
    ".webp": "image/webp",
    ".bmp": "image/bmp",
    ".mp4": "video/mp4",
    ".mp3": "audio/mpeg",
    ".wav": "audio/wav",
    ".html": "text/html",
    ".htm": "text/html",
    ".css": "text/css",
    ".js": "application/javascript",
    ".ts": "application/typescript",
    ".py": "text/x-python",
    ".md": "text/markdown",
    ".log": "text/plain",
};
/**
 * Detect MIME type from file extension.
 */
export function detectMimeType(filePath) {
    const ext = extname(filePath).toLowerCase();
    return MIME_TYPES[ext] || "application/octet-stream";
}
/**
 * Extract attachment GUID from the eTag returned by Microsoft Graph.
 * eTag format: `"{GUID},version"` → extracts the GUID portion.
 */
export function extractGuidFromETag(eTag) {
    const match = eTag.match(/\{([^}]+)\}/);
    if (match) {
        return match[1];
    }
    const [rawId] = eTag.split(",");
    return rawId.replace(/["{}]/g, "") || eTag;
}
/**
 * Read a local file and return its contents as a Buffer.
 */
export async function readLocalFile(filePath) {
    const buffer = await fs.readFile(filePath);
    return { buffer, size: buffer.length };
}
/**
 * Simple upload for files ≤ 4 MB.
 * PUT /drives/{driveId}/items/{parentItemId}:/{fileName}:/content
 */
async function simpleUpload(graphService, driveId, parentItemId, remotePath, fileBuffer, mimeType) {
    const client = await graphService.getClient();
    const response = (await client
        .api(`/drives/${driveId}/items/${parentItemId}:/${remotePath}:/content`)
        .header("Content-Type", mimeType)
        .put(fileBuffer));
    if (!response?.webUrl || !response?.eTag) {
        throw new Error("Upload failed: response did not contain webUrl/eTag");
    }
    return { webUrl: response.webUrl, eTag: response.eTag, id: response.id ?? "" };
}
/**
 * Upload session for files > 4 MB.
 * Creates a resumable upload session and sends the file in 3.2 MB chunks.
 */
async function uploadLargeFile(graphService, driveId, parentItemId, remotePath, fileBuffer) {
    const client = await graphService.getClient();
    const session = (await client
        .api(`/drives/${driveId}/items/${parentItemId}:/${remotePath}:/createUploadSession`)
        .post({
        item: {
            "@microsoft.graph.conflictBehavior": "rename",
        },
    }));
    if (!session?.uploadUrl) {
        throw new Error("Upload failed: upload session did not return uploadUrl");
    }
    const uploadUrl = session.uploadUrl;
    const fileSize = fileBuffer.length;
    let offset = 0;
    let lastResponse = null;
    while (offset < fileSize) {
        const chunkEnd = Math.min(offset + UPLOAD_CHUNK_SIZE, fileSize);
        const chunk = fileBuffer.subarray(offset, chunkEnd);
        const contentRange = `bytes ${offset}-${chunkEnd - 1}/${fileSize}`;
        lastResponse = await fetch(uploadUrl, {
            method: "PUT",
            headers: {
                "Content-Length": String(chunk.length),
                "Content-Range": contentRange,
            },
            body: new Uint8Array(chunk),
        });
        if (!lastResponse.ok) {
            const errorText = await lastResponse.text();
            throw new Error(`Upload chunk failed (${lastResponse.status}): ${errorText}`);
        }
        // Drain intermediate 202 response bodies to free resources
        if (lastResponse.status === 202) {
            await lastResponse.text();
        }
        offset = chunkEnd;
    }
    if (!lastResponse) {
        throw new Error("Upload failed: no response received");
    }
    const finalResult = await lastResponse.json();
    if (!finalResult?.webUrl || !finalResult?.eTag) {
        throw new Error("Upload failed: final response did not contain file metadata");
    }
    return { webUrl: finalResult.webUrl, eTag: finalResult.eTag, id: finalResult.id ?? "" };
}
/**
 * Upload a file to a Teams channel's SharePoint folder.
 */
export async function uploadFileToChannel(graphService, teamId, channelId, filePath, customFileName) {
    const { driveId, folderId: channelFolderId } = await getChannelDrive(graphService, teamId, channelId);
    const fileName = customFileName || basename(filePath);
    const mimeType = detectMimeType(filePath);
    const { buffer, size } = await readLocalFile(filePath);
    const encodedName = encodeURIComponent(fileName);
    const uploadResult = size <= SIMPLE_UPLOAD_MAX_SIZE
        ? await simpleUpload(graphService, driveId, channelFolderId, encodedName, buffer, mimeType)
        : await uploadLargeFile(graphService, driveId, channelFolderId, encodedName, buffer);
    return {
        webUrl: uploadResult.webUrl,
        attachmentId: extractGuidFromETag(uploadResult.eTag),
        fileName,
        fileSize: size,
        mimeType,
    };
}
/**
 * Upload a file to OneDrive's "Microsoft Teams Chat Files" folder for chat messages.
 */
export async function uploadFileToChat(graphService, filePath, customFileName) {
    const fileName = customFileName || basename(filePath);
    const mimeType = detectMimeType(filePath);
    const { buffer, size } = await readLocalFile(filePath);
    const driveId = await getMyDriveId(graphService);
    const remotePath = `${encodeURIComponent("Microsoft Teams Chat Files")}/${encodeURIComponent(fileName)}`;
    const uploadResult = size <= SIMPLE_UPLOAD_MAX_SIZE
        ? await simpleUpload(graphService, driveId, "root", remotePath, buffer, mimeType)
        : await uploadLargeFile(graphService, driveId, "root", remotePath, buffer);
    const attachmentId = extractGuidFromETag(uploadResult.eTag);
    // Chat message file attachments require a sharing link URL — the direct
    // webUrl from the upload causes "permission denied" for recipients.
    let contentUrl = uploadResult.webUrl;
    if (uploadResult.id) {
        contentUrl =
            (await createChatSharingLink(graphService, driveId, uploadResult.id)) ?? contentUrl;
    }
    return {
        webUrl: contentUrl,
        attachmentId,
        fileName,
        fileSize: size,
        mimeType,
    };
}
/**
 * Build the attachments array for a message that references an uploaded file.
 */
export function buildFileAttachment(uploadResult) {
    return [
        {
            id: uploadResult.attachmentId,
            contentType: "reference",
            contentUrl: uploadResult.webUrl,
            name: uploadResult.fileName,
        },
    ];
}
/**
 * Escape special HTML characters in plain text so it can be safely
 * embedded inside an HTML message body.
 */
export function escapeHtml(text) {
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}
/**
 * Format a file size in bytes to a human-readable string.
 */
export function formatFileSize(bytes) {
    if (bytes < 1024)
        return `${bytes} B`;
    if (bytes < 1024 * 1024)
        return `${(bytes / 1024).toFixed(1)} KB`;
    if (bytes < 1024 * 1024 * 1024)
        return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
    return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}
/** Resolve the current user's OneDrive id. */
async function getMyDriveId(graphService) {
    const client = await graphService.getClient();
    const driveResponse = (await client.api("/me/drive").get());
    if (!driveResponse?.id) {
        throw new Error("Failed to resolve user drive ID");
    }
    return driveResponse.id;
}
/** Resolve a channel's SharePoint drive and folder ids. */
async function getChannelDrive(graphService, teamId, channelId) {
    const client = await graphService.getClient();
    const filesFolder = (await client
        .api(`/teams/${teamId}/channels/${channelId}/filesFolder`)
        .get());
    if (!filesFolder?.parentReference?.driveId || !filesFolder?.id) {
        throw new Error("Failed to resolve channel drive/folder IDs");
    }
    return { driveId: filesFolder.parentReference.driveId, folderId: filesFolder.id };
}
/**
 * Create an organization-scoped sharing link for a drive item; falls back to
 * "users" scope when tenant policy blocks it. Returns undefined if both fail —
 * the caller can then fall back to the item's direct webUrl.
 */
export async function createChatSharingLink(graphService, driveId, itemId) {
    const client = await graphService.getClient();
    try {
        const linkResponse = (await client
            .api(`/drives/${driveId}/items/${itemId}/createLink`)
            .post({ type: "view", scope: "organization" }));
        if (linkResponse?.link?.webUrl) {
            return linkResponse.link.webUrl;
        }
    }
    catch (orgErr) {
        console.error(`[teams-mcp] createLink (organization) failed for item ${itemId}:`, orgErr instanceof Error ? orgErr.message : orgErr);
        try {
            const linkResponse = (await client
                .api(`/drives/${driveId}/items/${itemId}/createLink`)
                .post({ type: "view", scope: "users" }));
            if (linkResponse?.link?.webUrl) {
                return linkResponse.link.webUrl;
            }
        }
        catch (usersErr) {
            console.error(`[teams-mcp] createLink (users) also failed for item ${itemId}:`, usersErr instanceof Error ? usersErr.message : usersErr);
        }
    }
    return undefined;
}
/** Fetch an already-uploaded drive item and build a FileUploadResult for it. */
async function resolveDriveItem(graphService, driveId, itemId) {
    const client = await graphService.getClient();
    const item = (await client.api(`/drives/${driveId}/items/${itemId}`).get());
    if (!item?.webUrl || !item?.eTag || !item?.name) {
        throw new Error("Failed to resolve uploaded file metadata (webUrl/eTag/name)");
    }
    return {
        webUrl: item.webUrl,
        attachmentId: extractGuidFromETag(item.eTag),
        fileName: item.name,
        fileSize: item.size ?? 0,
        mimeType: item.file?.mimeType || detectMimeType(item.name),
    };
}
/**
 * Build a FileUploadResult for a file already uploaded to the user's OneDrive
 * (e.g. via create_file_upload_session), including the sharing link chats need.
 */
export async function resolveChatDriveItem(graphService, driveItemId) {
    const driveId = await getMyDriveId(graphService);
    const result = await resolveDriveItem(graphService, driveId, driveItemId);
    const sharingLink = await createChatSharingLink(graphService, driveId, driveItemId);
    return sharingLink ? { ...result, webUrl: sharingLink } : result;
}
/**
 * Build a FileUploadResult for a file already uploaded to a channel's
 * SharePoint folder (e.g. via create_file_upload_session).
 */
export async function resolveChannelDriveItem(graphService, teamId, channelId, driveItemId) {
    const { driveId } = await getChannelDrive(graphService, teamId, channelId);
    return resolveDriveItem(graphService, driveId, driveItemId);
}
async function createUploadSession(graphService, driveId, parentItemId, remotePath) {
    const client = await graphService.getClient();
    const session = (await client
        .api(`/drives/${driveId}/items/${parentItemId}:/${remotePath}:/createUploadSession`)
        .post({
        item: { "@microsoft.graph.conflictBehavior": "rename" },
    }));
    if (!session?.uploadUrl) {
        throw new Error("Failed to create upload session: no uploadUrl returned");
    }
    return { uploadUrl: session.uploadUrl, expirationDateTime: session.expirationDateTime };
}
/** Create a resumable upload session in the user's "Microsoft Teams Chat Files" folder. */
export async function createUploadSessionForChat(graphService, fileName) {
    const driveId = await getMyDriveId(graphService);
    const remotePath = `${encodeURIComponent("Microsoft Teams Chat Files")}/${encodeURIComponent(fileName)}`;
    return createUploadSession(graphService, driveId, "root", remotePath);
}
/** Create a resumable upload session in a channel's SharePoint folder. */
export async function createUploadSessionForChannel(graphService, teamId, channelId, fileName) {
    const { driveId, folderId } = await getChannelDrive(graphService, teamId, channelId);
    return createUploadSession(graphService, driveId, folderId, encodeURIComponent(fileName));
}
/** Encode a SharePoint/OneDrive URL for the Graph Shares API ("u!" base64url format). */
export function encodeShareUrl(url) {
    return `u!${Buffer.from(url)
        .toString("base64")
        .replace(/\+/g, "-")
        .replace(/\//g, "_")
        .replace(/=+$/, "")}`;
}
//# sourceMappingURL=file-upload.js.map