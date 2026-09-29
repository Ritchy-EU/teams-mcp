/**
 * Request header that makes Graph return `messageType: "systemEventMessage"`
 * instead of `unknownFutureValue` for system event messages.
 * See https://learn.microsoft.com/graph/api/resources/chatmessage (messageType).
 */
export const INCLUDE_UNKNOWN_ENUM_MEMBERS_HEADER = [
    "Prefer",
    "include-unknown-enum-members",
];
const ODATA_TYPE_PREFIX = "#microsoft.graph.";
const CALL_STARTED = "callStartedEventMessageDetail";
const CALL_ENDED = "callEndedEventMessageDetail";
/**
 * Returns true when the message is a Teams system event (member added, call ended, ...)
 * rather than a user-authored message.
 * Without the `Prefer: include-unknown-enum-members` header Graph reports these as
 * `unknownFutureValue`, so the presence of `eventDetail` is the reliable signal.
 */
export function isSystemEventMessage(message) {
    return message.eventDetail != null || message.messageType === "systemEventMessage";
}
/**
 * Strips the "#microsoft.graph." prefix from an eventDetail @odata.type.
 */
export function eventDetailType(eventDetail) {
    const odataType = eventDetail?.["@odata.type"];
    if (!odataType)
        return "unknown";
    return odataType.startsWith(ODATA_TYPE_PREFIX)
        ? odataType.slice(ODATA_TYPE_PREFIX.length)
        : odataType.replace(/^#/, "");
}
/**
 * Returns true when the message is a callStarted or callEnded system event.
 */
export function isCallEventMessage(message) {
    const type = eventDetailType(message.eventDetail);
    return type === CALL_STARTED || type === CALL_ENDED;
}
/**
 * Parses an ISO 8601 duration (as returned by Graph, e.g. "PT39M29S" or "P1DT2H")
 * into whole seconds. Returns undefined for missing or unparseable input.
 * Years and months are not supported (Graph never emits them for call durations).
 */
export function parseIsoDurationSeconds(duration) {
    if (!duration)
        return undefined;
    const match = /^(-)?P(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(duration.trim());
    if (!match)
        return undefined;
    const [, sign, days, hours, minutes, seconds] = match;
    if (!days && !hours && !minutes && !seconds)
        return undefined;
    const total = Number(days ?? 0) * 86400 +
        Number(hours ?? 0) * 3600 +
        Number(minutes ?? 0) * 60 +
        Number(seconds ?? 0);
    const rounded = Math.round(total);
    return sign ? -rounded : rounded;
}
function summarizeIdentity(identity) {
    const user = identity?.user;
    if (!user)
        return undefined;
    return {
        id: user.id ?? undefined,
        // Graph sends "" for the initiator's displayName; treat it as absent.
        displayName: user.displayName || undefined,
    };
}
/**
 * Flattens a chatMessage.eventDetail into an EventDetailSummary.
 * Call events get named fields; other event types are passed through in `raw`.
 */
export function summarizeEventDetail(eventDetail) {
    if (!eventDetail)
        return undefined;
    const type = eventDetailType(eventDetail);
    if (type === CALL_STARTED) {
        const detail = eventDetail;
        return {
            type,
            callId: detail.callId ?? undefined,
            callEventType: detail.callEventType ?? undefined,
            initiator: summarizeIdentity(detail.initiator),
        };
    }
    if (type === CALL_ENDED) {
        const detail = eventDetail;
        const participants = (detail.callParticipants ?? [])
            .map((p) => summarizeIdentity(p.participant))
            .filter((p) => p !== undefined);
        return {
            type,
            callId: detail.callId ?? undefined,
            callEventType: detail.callEventType ?? undefined,
            callDuration: detail.callDuration ?? undefined,
            callDurationSeconds: parseIsoDurationSeconds(detail.callDuration),
            initiator: summarizeIdentity(detail.initiator),
            callParticipants: participants,
        };
    }
    return { type, raw: eventDetail };
}
/**
 * Reconstructs calls from callStarted / callEnded system messages, joined on callId.
 * - startDateTime = createdDateTime of callStarted, endDateTime = createdDateTime of callEnded.
 * - When callStarted is missing but a duration is known, start = end - duration and
 *   `startEstimated` is set.
 * Result is sorted newest first by (start ?? end).
 */
export function stitchCalls(messages) {
    const byCallId = new Map();
    for (const message of messages) {
        const summary = summarizeEventDetail(message.eventDetail);
        if (!summary?.callId)
            continue;
        if (summary.type !== CALL_STARTED && summary.type !== CALL_ENDED)
            continue;
        const call = byCallId.get(summary.callId) ?? { callId: summary.callId };
        call.callEventType = call.callEventType ?? summary.callEventType;
        if (summary.type === CALL_STARTED) {
            call.startDateTime = message.createdDateTime ?? undefined;
            call.startEstimated = undefined;
            call.initiator = summary.initiator ?? call.initiator;
        }
        else {
            call.endDateTime = message.createdDateTime ?? undefined;
            call.durationSeconds = summary.callDurationSeconds;
            call.initiator = call.initiator ?? summary.initiator;
            if (summary.callParticipants?.length)
                call.participants = summary.callParticipants;
        }
        byCallId.set(summary.callId, call);
    }
    for (const call of byCallId.values()) {
        // Graph omits the initiator's displayName on call events; borrow it from the participant list.
        if (call.initiator?.id && !call.initiator.displayName) {
            const match = call.participants?.find((p) => p.id === call.initiator?.id);
            if (match?.displayName)
                call.initiator = { ...call.initiator, displayName: match.displayName };
        }
        if (!call.startDateTime && call.endDateTime && call.durationSeconds !== undefined) {
            const end = new Date(call.endDateTime).getTime();
            if (!Number.isNaN(end)) {
                call.startDateTime = new Date(end - call.durationSeconds * 1000).toISOString();
                call.startEstimated = true;
            }
        }
    }
    const sortKey = (c) => new Date(c.startDateTime ?? c.endDateTime ?? 0).getTime();
    return [...byCallId.values()].sort((a, b) => sortKey(b) - sortKey(a));
}
//# sourceMappingURL=events.js.map