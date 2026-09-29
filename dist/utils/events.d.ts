import type { CallSummary, ChatMessage, EventDetailSummary, EventMessageDetail } from "../types/graph.js";
/**
 * Request header that makes Graph return `messageType: "systemEventMessage"`
 * instead of `unknownFutureValue` for system event messages.
 * See https://learn.microsoft.com/graph/api/resources/chatmessage (messageType).
 */
export declare const INCLUDE_UNKNOWN_ENUM_MEMBERS_HEADER: readonly ["Prefer", "include-unknown-enum-members"];
/**
 * Returns true when the message is a Teams system event (member added, call ended, ...)
 * rather than a user-authored message.
 * Without the `Prefer: include-unknown-enum-members` header Graph reports these as
 * `unknownFutureValue`, so the presence of `eventDetail` is the reliable signal.
 */
export declare function isSystemEventMessage(message: ChatMessage): boolean;
/**
 * Strips the "#microsoft.graph." prefix from an eventDetail @odata.type.
 */
export declare function eventDetailType(eventDetail: EventMessageDetail | null | undefined): string;
/**
 * Returns true when the message is a callStarted or callEnded system event.
 */
export declare function isCallEventMessage(message: ChatMessage): boolean;
/**
 * Parses an ISO 8601 duration (as returned by Graph, e.g. "PT39M29S" or "P1DT2H")
 * into whole seconds. Returns undefined for missing or unparseable input.
 * Years and months are not supported (Graph never emits them for call durations).
 */
export declare function parseIsoDurationSeconds(duration: string | null | undefined): number | undefined;
/**
 * Flattens a chatMessage.eventDetail into an EventDetailSummary.
 * Call events get named fields; other event types are passed through in `raw`.
 */
export declare function summarizeEventDetail(eventDetail: EventMessageDetail | null | undefined): EventDetailSummary | undefined;
/**
 * Reconstructs calls from callStarted / callEnded system messages, joined on callId.
 * - startDateTime = createdDateTime of callStarted, endDateTime = createdDateTime of callEnded.
 * - When callStarted is missing but a duration is known, start = end - duration and
 *   `startEstimated` is set.
 * Result is sorted newest first by (start ?? end).
 */
export declare function stitchCalls(messages: ChatMessage[]): CallSummary[];
//# sourceMappingURL=events.d.ts.map