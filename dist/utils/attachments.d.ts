import type { ChatMessageAttachment } from "@microsoft/microsoft-graph-types";
import type { IGraphService } from "../services/graph.js";
import type { AttachmentSummary } from "../types/graph.js";
export interface ImageAttachment {
    id: string;
    contentType: string;
    contentUrl?: string;
    name?: string;
    thumbnailUrl?: string;
}
export interface HostedContent {
    "@microsoft.graph.temporaryId": string;
    contentBytes: string;
    contentType: string;
}
/**
 * Upload image as hosted content for Teams messages
 * This creates a temporary hosted content that can be referenced in message attachments
 */
export declare function uploadImageAsHostedContent(graphService: IGraphService, teamId: string, channelId: string, imageData: Buffer | string, contentType: string, fileName?: string): Promise<{
    hostedContentId: string;
    attachment: ImageAttachment;
} | null>;
/**
 * Validate image content type
 */
export declare function isValidImageType(contentType: string): boolean;
/**
 * Get file extension from MIME type
 */
export declare function getFileExtensionFromMimeType(mimeType: string): string;
/**
 * Convert image URL to base64 for upload
 */
export declare function imageUrlToBase64(imageUrl: string): Promise<{
    data: string;
    contentType: string;
} | null>;
/**
 * Extracts a minimal attachment summary from Graph API ChatMessageAttachment array.
 * Returns undefined if there are no meaningful attachments to report.
 */
export declare function extractAttachmentSummaries(attachments: ChatMessageAttachment[] | null | undefined): AttachmentSummary[] | undefined;
//# sourceMappingURL=attachments.d.ts.map