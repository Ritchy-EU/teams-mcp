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
//# sourceMappingURL=file-upload.d.ts.map