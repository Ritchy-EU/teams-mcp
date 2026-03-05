import type { IGraphService } from "../services/graph.js";
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
//# sourceMappingURL=attachments.d.ts.map