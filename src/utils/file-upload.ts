import { promises as fs } from "node:fs";
import { basename, extname } from "node:path";
import type { IGraphService } from "../services/graph.js";

/** Simple upload threshold (4 MB) */
const SIMPLE_UPLOAD_MAX_SIZE = 4 * 1024 * 1024;

/** Upload session chunk size — must be a multiple of 320 KiB */
const UPLOAD_CHUNK_SIZE = 320 * 1024 * 10; // 3.2 MB

/** Extension → MIME type map for common file types */
const MIME_TYPES: Record<string, string> = {
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

export interface FileUploadResult {
  webUrl: string;
  attachmentId: string;
  fileName: string;
  fileSize: number;
  mimeType: string;
  driveId?: string | undefined;
  itemId?: string | undefined;
}

/** Graph API response from a DriveItem upload (simple PUT or final chunk). */
type DriveItemUploadResponse = {
  id?: string;
  webUrl?: string;
  eTag?: string;
};

/** Graph API response when creating a resumable upload session. */
type UploadSessionResponse = {
  uploadUrl?: string;
};

/** Graph API response for the channel filesFolder endpoint. */
type ChannelFilesFolderResponse = {
  id?: string;
  parentReference?: { driveId?: string };
};

/** Graph API response from the createLink endpoint. */
type CreateLinkResponse = {
  link?: {
    webUrl?: string;
  };
};

/**
 * Detect MIME type from file extension.
 */
export function detectMimeType(filePath: string): string {
  const ext = extname(filePath).toLowerCase();
  return MIME_TYPES[ext] || "application/octet-stream";
}

/**
 * Extract attachment GUID from the eTag returned by Microsoft Graph.
 * eTag format: `"{GUID},version"` → extracts the GUID portion.
 */
export function extractGuidFromETag(eTag: string): string {
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
export async function readLocalFile(filePath: string): Promise<{ buffer: Buffer; size: number }> {
  const buffer = await fs.readFile(filePath);
  return { buffer, size: buffer.length };
}

/**
 * Simple upload for files ≤ 4 MB.
 * PUT /drives/{driveId}/items/{parentItemId}:/{fileName}:/content
 */
async function simpleUpload(
  graphService: IGraphService,
  driveId: string,
  parentItemId: string,
  remotePath: string,
  fileBuffer: Buffer,
  mimeType: string
): Promise<{ webUrl: string; eTag: string; id: string }> {
  const client = await graphService.getClient();
  const response = (await client
    .api(`/drives/${driveId}/items/${parentItemId}:/${remotePath}:/content`)
    .header("Content-Type", mimeType)
    .put(fileBuffer)) as DriveItemUploadResponse;
  if (!response?.webUrl || !response?.eTag) {
    throw new Error("Upload failed: response did not contain webUrl/eTag");
  }
  return { webUrl: response.webUrl, eTag: response.eTag, id: response.id ?? "" };
}

/**
 * Upload session for files > 4 MB.
 * Creates a resumable upload session and sends the file in 3.2 MB chunks.
 */
async function uploadLargeFile(
  graphService: IGraphService,
  driveId: string,
  parentItemId: string,
  remotePath: string,
  fileBuffer: Buffer
): Promise<{ webUrl: string; eTag: string; id: string }> {
  const client = await graphService.getClient();

  const session = (await client
    .api(`/drives/${driveId}/items/${parentItemId}:/${remotePath}:/createUploadSession`)
    .post({
      item: {
        "@microsoft.graph.conflictBehavior": "rename",
      },
    })) as UploadSessionResponse;

  if (!session?.uploadUrl) {
    throw new Error("Upload failed: upload session did not return uploadUrl");
  }
  const uploadUrl: string = session.uploadUrl;
  const fileSize = fileBuffer.length;

  let offset = 0;
  let lastResponse: Response | null = null;

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
export async function uploadFileToChannel(
  graphService: IGraphService,
  teamId: string,
  channelId: string,
  filePath: string,
  customFileName?: string
): Promise<FileUploadResult> {
  const { driveId, folderId: channelFolderId } = await getChannelDrive(
    graphService,
    teamId,
    channelId
  );

  const fileName = customFileName || basename(filePath);
  const mimeType = detectMimeType(filePath);
  const { buffer, size } = await readLocalFile(filePath);

  const encodedName = encodeURIComponent(fileName);
  const uploadResult =
    size <= SIMPLE_UPLOAD_MAX_SIZE
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
export async function uploadFileToChat(
  graphService: IGraphService,
  filePath: string,
  customFileName?: string
): Promise<FileUploadResult> {
  const fileName = customFileName || basename(filePath);
  const mimeType = detectMimeType(filePath);
  const { buffer, size } = await readLocalFile(filePath);

  const driveId = await getMyDriveId(graphService);

  const remotePath = `${encodeURIComponent("Microsoft Teams Chat Files")}/${encodeURIComponent(fileName)}`;
  const uploadResult =
    size <= SIMPLE_UPLOAD_MAX_SIZE
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
    driveId,
    itemId: uploadResult.id || undefined,
  };
}

/**
 * Build the attachments array for a message that references an uploaded file.
 */
export function buildFileAttachment(uploadResult: FileUploadResult): Array<{
  id: string;
  contentType: string;
  contentUrl: string;
  name: string;
}> {
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
export function escapeHtml(text: string): string {
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
export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

/** Basic driveItem metadata returned by Graph for an uploaded file. */
type DriveItemMetadata = {
  id?: string;
  name?: string;
  size?: number;
  webUrl?: string;
  eTag?: string;
  file?: { mimeType?: string };
};

/** Resolve the current user's OneDrive id. */
async function getMyDriveId(graphService: IGraphService): Promise<string> {
  const client = await graphService.getClient();
  const driveResponse = (await client.api("/me/drive").get()) as { id?: string };
  if (!driveResponse?.id) {
    throw new Error("Failed to resolve user drive ID");
  }
  return driveResponse.id;
}

/** Resolve a channel's SharePoint drive and folder ids. */
async function getChannelDrive(
  graphService: IGraphService,
  teamId: string,
  channelId: string
): Promise<{ driveId: string; folderId: string }> {
  const client = await graphService.getClient();
  const filesFolder = (await client
    .api(`/teams/${teamId}/channels/${channelId}/filesFolder`)
    .get()) as ChannelFilesFolderResponse;
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
export async function createChatSharingLink(
  graphService: IGraphService,
  driveId: string,
  itemId: string
): Promise<string | undefined> {
  const client = await graphService.getClient();
  try {
    const linkResponse = (await client
      .api(`/drives/${driveId}/items/${itemId}/createLink`)
      .post({ type: "view", scope: "organization" })) as CreateLinkResponse;
    if (linkResponse?.link?.webUrl) {
      return linkResponse.link.webUrl;
    }
  } catch (orgErr: unknown) {
    console.error(
      `[teams-mcp] createLink (organization) failed for item ${itemId}:`,
      orgErr instanceof Error ? orgErr.message : orgErr
    );
    try {
      const linkResponse = (await client
        .api(`/drives/${driveId}/items/${itemId}/createLink`)
        .post({ type: "view", scope: "users" })) as CreateLinkResponse;
      if (linkResponse?.link?.webUrl) {
        return linkResponse.link.webUrl;
      }
    } catch (usersErr: unknown) {
      console.error(
        `[teams-mcp] createLink (users) also failed for item ${itemId}:`,
        usersErr instanceof Error ? usersErr.message : usersErr
      );
    }
  }
  return undefined;
}

/** Graph API response for a chat members listing (fields used for access grants). */
type ChatMembersResponse = {
  value?: Array<{ userId?: string; email?: string }>;
};

/**
 * Grant every other chat member direct read access to a drive item.
 * The Teams client does this synchronously when a user shares a file in a chat;
 * without it, recipients cannot download the file via Graph until they open it
 * in the Teams client once. sendInvitation:false adds the permission silently.
 * Returns the number of members granted access (0 when the chat has no other members).
 */
export async function grantChatMembersAccess(
  graphService: IGraphService,
  chatId: string,
  driveId: string,
  itemId: string
): Promise<number> {
  const client = await graphService.getClient();
  const me = (await client.api("/me").get()) as { id?: string };
  const membersResponse = (await client
    .api(`/chats/${chatId}/members`)
    .get()) as ChatMembersResponse;

  const recipients: Array<{ email: string } | { objectId: string }> = [];
  const seen = new Set<string>();
  for (const member of membersResponse?.value ?? []) {
    const userId = member.userId;
    if (!userId || userId === me?.id || seen.has(userId)) {
      continue;
    }
    seen.add(userId);
    recipients.push(member.email ? { email: member.email } : { objectId: userId });
  }

  if (recipients.length === 0) {
    return 0;
  }

  await client.api(`/drives/${driveId}/items/${itemId}/invite`).post({
    recipients,
    requireSignIn: true,
    sendInvitation: false,
    roles: ["read"],
  });
  return recipients.length;
}

/** Fetch an already-uploaded drive item and build a FileUploadResult for it. */
async function resolveDriveItem(
  graphService: IGraphService,
  driveId: string,
  itemId: string
): Promise<FileUploadResult> {
  const client = await graphService.getClient();
  const item = (await client.api(`/drives/${driveId}/items/${itemId}`).get()) as DriveItemMetadata;
  if (!item?.webUrl || !item?.eTag || !item?.name) {
    throw new Error("Failed to resolve uploaded file metadata (webUrl/eTag/name)");
  }
  return {
    webUrl: item.webUrl,
    attachmentId: extractGuidFromETag(item.eTag),
    fileName: item.name,
    fileSize: item.size ?? 0,
    mimeType: item.file?.mimeType || detectMimeType(item.name),
    driveId,
    itemId: item.id ?? itemId,
  };
}

/**
 * Build a FileUploadResult for a file already uploaded to the user's OneDrive
 * (e.g. via create_file_upload_session), including the sharing link chats need.
 */
export async function resolveChatDriveItem(
  graphService: IGraphService,
  driveItemId: string
): Promise<FileUploadResult> {
  const driveId = await getMyDriveId(graphService);
  const result = await resolveDriveItem(graphService, driveId, driveItemId);
  const sharingLink = await createChatSharingLink(graphService, driveId, driveItemId);
  return sharingLink ? { ...result, webUrl: sharingLink } : result;
}

/**
 * Build a FileUploadResult for a file already uploaded to a channel's
 * SharePoint folder (e.g. via create_file_upload_session).
 */
export async function resolveChannelDriveItem(
  graphService: IGraphService,
  teamId: string,
  channelId: string,
  driveItemId: string
): Promise<FileUploadResult> {
  const { driveId } = await getChannelDrive(graphService, teamId, channelId);
  return resolveDriveItem(graphService, driveId, driveItemId);
}

export interface UploadSessionInfo {
  uploadUrl: string;
  expirationDateTime?: string | undefined;
}

async function createUploadSession(
  graphService: IGraphService,
  driveId: string,
  parentItemId: string,
  remotePath: string
): Promise<UploadSessionInfo> {
  const client = await graphService.getClient();
  const session = (await client
    .api(`/drives/${driveId}/items/${parentItemId}:/${remotePath}:/createUploadSession`)
    .post({
      item: { "@microsoft.graph.conflictBehavior": "rename" },
    })) as UploadSessionResponse & { expirationDateTime?: string };
  if (!session?.uploadUrl) {
    throw new Error("Failed to create upload session: no uploadUrl returned");
  }
  return { uploadUrl: session.uploadUrl, expirationDateTime: session.expirationDateTime };
}

/** Create a resumable upload session in the user's "Microsoft Teams Chat Files" folder. */
export async function createUploadSessionForChat(
  graphService: IGraphService,
  fileName: string
): Promise<UploadSessionInfo> {
  const driveId = await getMyDriveId(graphService);
  const remotePath = `${encodeURIComponent("Microsoft Teams Chat Files")}/${encodeURIComponent(fileName)}`;
  return createUploadSession(graphService, driveId, "root", remotePath);
}

/** Create a resumable upload session in a channel's SharePoint folder. */
export async function createUploadSessionForChannel(
  graphService: IGraphService,
  teamId: string,
  channelId: string,
  fileName: string
): Promise<UploadSessionInfo> {
  const { driveId, folderId } = await getChannelDrive(graphService, teamId, channelId);
  return createUploadSession(graphService, driveId, folderId, encodeURIComponent(fileName));
}

/** Encode a SharePoint/OneDrive URL for the Graph Shares API ("u!" base64url format). */
export function encodeShareUrl(url: string): string {
  return `u!${Buffer.from(url)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "")}`;
}
