import type { IGraphService } from "../services/graph.js";
export interface FileUploadResult {
    webUrl: string;
    attachmentId: string;
    fileName: string;
    fileSize: number;
    mimeType: string;
}
/**
 * Detect MIME type from file extension.
 */
export declare function detectMimeType(filePath: string): string;
/**
 * Extract attachment GUID from the eTag returned by Microsoft Graph.
 * eTag format: `"{GUID},version"` → extracts the GUID portion.
 */
export declare function extractGuidFromETag(eTag: string): string;
/**
 * Read a local file and return its contents as a Buffer.
 */
export declare function readLocalFile(filePath: string): Promise<{
    buffer: Buffer;
    size: number;
}>;
/**
 * Upload a file to a Teams channel's SharePoint folder.
 */
export declare function uploadFileToChannel(graphService: IGraphService, teamId: string, channelId: string, filePath: string, customFileName?: string): Promise<FileUploadResult>;
/**
 * Upload a file to OneDrive's "Microsoft Teams Chat Files" folder for chat messages.
 */
export declare function uploadFileToChat(graphService: IGraphService, filePath: string, customFileName?: string): Promise<FileUploadResult>;
/**
 * Build the attachments array for a message that references an uploaded file.
 */
export declare function buildFileAttachment(uploadResult: FileUploadResult): Array<{
    id: string;
    contentType: string;
    contentUrl: string;
    name: string;
}>;
/**
 * Escape special HTML characters in plain text so it can be safely
 * embedded inside an HTML message body.
 */
export declare function escapeHtml(text: string): string;
/**
 * Format a file size in bytes to a human-readable string.
 */
export declare function formatFileSize(bytes: number): string;
/**
 * Create an organization-scoped sharing link for a drive item; falls back to
 * "users" scope when tenant policy blocks it. Returns undefined if both fail —
 * the caller can then fall back to the item's direct webUrl.
 */
export declare function createChatSharingLink(graphService: IGraphService, driveId: string, itemId: string): Promise<string | undefined>;
/**
 * Build a FileUploadResult for a file already uploaded to the user's OneDrive
 * (e.g. via create_file_upload_session), including the sharing link chats need.
 */
export declare function resolveChatDriveItem(graphService: IGraphService, driveItemId: string): Promise<FileUploadResult>;
/**
 * Build a FileUploadResult for a file already uploaded to a channel's
 * SharePoint folder (e.g. via create_file_upload_session).
 */
export declare function resolveChannelDriveItem(graphService: IGraphService, teamId: string, channelId: string, driveItemId: string): Promise<FileUploadResult>;
export interface UploadSessionInfo {
    uploadUrl: string;
    expirationDateTime?: string | undefined;
}
/** Create a resumable upload session in the user's "Microsoft Teams Chat Files" folder. */
export declare function createUploadSessionForChat(graphService: IGraphService, fileName: string): Promise<UploadSessionInfo>;
/** Create a resumable upload session in a channel's SharePoint folder. */
export declare function createUploadSessionForChannel(graphService: IGraphService, teamId: string, channelId: string, fileName: string): Promise<UploadSessionInfo>;
/** Encode a SharePoint/OneDrive URL for the Graph Shares API ("u!" base64url format). */
export declare function encodeShareUrl(url: string): string;
//# sourceMappingURL=file-upload.d.ts.map