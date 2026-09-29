import type {
  AadUserConversationMember,
  CallEndedEventMessageDetail,
  CallStartedEventMessageDetail,
  Channel,
  ChannelMembershipType,
  Chat,
  ChatMessage,
  ChatMessageAttachment,
  ChatMessageImportance,
  ChatMessageInfo,
  ChatMessageReaction,
  ChatType,
  ConversationMember,
  DirectoryObject,
  EventMessageDetail,
  IdentitySet,
  NullableOption,
  Team,
  TeamSpecialization,
  TeamsAppInstallation,
  TeamVisibilityType,
  User,
} from "@microsoft/microsoft-graph-types";

// Re-export Microsoft Graph types we use
export type {
  AadUserConversationMember,
  User,
  Chat,
  Team,
  Channel,
  ChatMessage,
  ChatMessageAttachment,
  ChatMessageReaction,
  ConversationMember,
  DirectoryObject,
  TeamsAppInstallation,
  ChatMessageInfo,
  ChannelMembershipType,
  ChatType,
  ChatMessageImportance,
  TeamSpecialization,
  TeamVisibilityType,
  NullableOption,
  EventMessageDetail,
  CallStartedEventMessageDetail,
  CallEndedEventMessageDetail,
  IdentitySet,
};

// Custom types for our responses
export interface GraphApiResponse<T> {
  value?: T[];
  "@odata.count"?: number;
  "@odata.nextLink"?: string;
}

export interface GraphError {
  code: string;
  message: string;
  innerError?: {
    code?: string;
    message?: string;
    "request-id"?: string;
    date?: string;
  };
}

// Simplified types for our API responses - all properties are optional to handle Graph API variability
export interface UserSummary {
  id?: string | undefined;
  displayName?: NullableOption<string> | undefined;
  userPrincipalName?: NullableOption<string> | undefined;
  mail?: NullableOption<string> | undefined;
  jobTitle?: NullableOption<string> | undefined;
  department?: NullableOption<string> | undefined;
  officeLocation?: NullableOption<string> | undefined;
}

export interface TeamSummary {
  id?: string | undefined;
  displayName?: NullableOption<string> | undefined;
  description?: NullableOption<string> | undefined;
  isArchived?: NullableOption<boolean> | undefined;
}

export interface ChannelSummary {
  id?: string | undefined;
  displayName?: string | undefined;
  description?: NullableOption<string> | undefined;
  membershipType?: NullableOption<ChannelMembershipType> | undefined;
}

export interface ChatSummary {
  id?: string | undefined;
  topic?: NullableOption<string> | undefined;
  chatType?: ChatType | undefined;
  memberCount?: number | undefined;
}

export interface AttachmentSummary {
  id?: string | undefined;
  name?: string | undefined;
  contentType?: string | undefined;
  contentUrl?: string | undefined;
  thumbnailUrl?: string | undefined;
}

export interface ReactionSummary {
  reactionType?: string | undefined;
  displayName?: NullableOption<string> | undefined;
  createdDateTime?: string | undefined;
  user?: { id?: string | undefined; displayName?: string | undefined } | undefined;
}

export interface IdentitySummary {
  id?: string | undefined;
  displayName?: string | undefined;
}

/**
 * Flattened view of a chatMessage.eventDetail (system event message).
 * Call events (callStarted / callEnded) are mapped to named fields;
 * every other event type is passed through unchanged in `raw`.
 */
export interface EventDetailSummary {
  /** "@odata.type" without the "#microsoft.graph." prefix, e.g. "callEndedEventMessageDetail". */
  type: string;
  callId?: string | undefined;
  /** call | meeting | screenShare */
  callEventType?: string | undefined;
  /** ISO 8601 duration as returned by Graph, e.g. "PT39M29S". Only on callEnded. */
  callDuration?: string | undefined;
  callDurationSeconds?: number | undefined;
  initiator?: IdentitySummary | undefined;
  callParticipants?: IdentitySummary[] | undefined;
  /** Non-call events (membersAdded, chatRenamed, ...) are passed through as-is. */
  raw?: unknown;
}

/**
 * A call reconstructed from its callStarted / callEnded system messages.
 */
export interface CallSummary {
  callId: string;
  callEventType?: string | undefined;
  startDateTime?: string | undefined;
  /** True when no callStarted event was in the fetched range and start was derived from end - duration. */
  startEstimated?: boolean | undefined;
  endDateTime?: string | undefined;
  durationSeconds?: number | undefined;
  initiator?: IdentitySummary | undefined;
  participants?: IdentitySummary[] | undefined;
}

export interface MessageSummary {
  id?: string | undefined;
  content?: NullableOption<string> | undefined;
  from?: NullableOption<string> | undefined;
  fromId?: string | undefined;
  createdDateTime?: NullableOption<string> | undefined;
  lastEditedDateTime?: NullableOption<string> | undefined;
  deletedDateTime?: NullableOption<string> | undefined;
  messageType?: string | undefined;
  importance?: ChatMessageImportance | undefined;
  attachments?: AttachmentSummary[] | undefined;
  reactions?: ReactionSummary[] | undefined;
  eventDetail?: EventDetailSummary | undefined;
}

export interface MemberSummary {
  id?: string | undefined;
  displayName?: NullableOption<string> | undefined;
  roles?: NullableOption<string[]> | undefined;
}

// Create chat payload
export interface CreateChatPayload {
  chatType: "oneOnOne" | "group";
  members: ConversationMember[];
  topic?: string;
}

// Send message payload
export interface SendMessagePayload {
  body: {
    content: string;
    contentType: "text" | "html";
  };
  importance?: ChatMessageImportance;
}

// New types for search functionality
export interface SearchRequest {
  entityTypes: string[];
  query: {
    queryString: string;
  };
  from?: number;
  size?: number;
  enableTopResults?: boolean;
}

export interface SearchResponse {
  value: SearchResult[];
}

export interface SearchResult {
  searchTerms: string[];
  hitsContainers: SearchHitsContainer[];
}

export interface SearchHitsContainer {
  hits: SearchHit[];
  total: number;
  moreResultsAvailable: boolean;
}

export interface SearchHit {
  hitId: string;
  rank: number;
  summary: string;
  resource: {
    "@odata.type": string;
    id: string;
    createdDateTime?: string;
    lastModifiedDateTime?: string;
    from?: {
      user?: {
        displayName?: string;
        id?: string;
      };
    };
    body?: {
      content?: string;
      contentType?: string;
    };
    subject?: string;
    importance?: string;
    webLink?: string;
    chatId?: string;
    channelIdentity?: {
      teamId?: string;
      channelId?: string;
    };
  };
}
