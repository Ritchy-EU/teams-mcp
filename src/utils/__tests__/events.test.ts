import { describe, expect, it } from "vitest";
import type { ChatMessage } from "../../types/graph.js";
import {
  eventDetailType,
  INCLUDE_UNKNOWN_ENUM_MEMBERS_HEADER,
  isCallEventMessage,
  isSystemEventMessage,
  parseIsoDurationSeconds,
  stitchCalls,
  summarizeEventDetail,
} from "../events.js";

const callStarted = (
  id: string,
  createdDateTime: string,
  callId = "call-1",
  callEventType = "call"
): ChatMessage =>
  ({
    id,
    messageType: "systemEventMessage",
    createdDateTime,
    body: { contentType: "html", content: "<systemEventMessage/>" },
    eventDetail: {
      "@odata.type": "#microsoft.graph.callStartedEventMessageDetail",
      callId,
      callEventType,
      initiator: {
        application: null,
        device: null,
        user: { id: "user-a", displayName: "Alice", userIdentityType: "aadUser" },
      },
    },
  }) as unknown as ChatMessage;

const callEnded = (
  id: string,
  createdDateTime: string,
  callDuration: string,
  callId = "call-1",
  callEventType = "call"
): ChatMessage =>
  ({
    id,
    messageType: "systemEventMessage",
    createdDateTime,
    body: { contentType: "html", content: "<systemEventMessage/>" },
    eventDetail: {
      "@odata.type": "#microsoft.graph.callEndedEventMessageDetail",
      callId,
      callDuration,
      callEventType,
      callParticipants: [
        { participant: { user: { id: "user-a", displayName: "Alice" } } },
        { participant: { user: { id: "user-b", displayName: "Bob" } } },
        { participant: { application: { id: "bot-1", displayName: "Bot" } } },
      ],
      initiator: { user: { id: "user-a", displayName: null } },
    },
  }) as unknown as ChatMessage;

describe("events utilities", () => {
  it("exposes the Prefer header pair", () => {
    expect(INCLUDE_UNKNOWN_ENUM_MEMBERS_HEADER).toEqual(["Prefer", "include-unknown-enum-members"]);
  });

  describe("parseIsoDurationSeconds", () => {
    it.each([
      ["PT59S", 59],
      ["PT39M29S", 39 * 60 + 29],
      ["PT1H2M3S", 3723],
      ["PT2H", 7200],
      ["P1DT1S", 86401],
      ["PT40M19.26S", 2419],
      ["PT0S", 0],
      ["-PT5S", -5],
    ])("parses %s", (input, expected) => {
      expect(parseIsoDurationSeconds(input)).toBe(expected);
    });

    it.each([
      undefined,
      null,
      "",
      "P",
      "PT",
      "garbage",
      "P1Y",
    ])("returns undefined for %s", (input) => {
      expect(parseIsoDurationSeconds(input)).toBeUndefined();
    });
  });

  describe("eventDetailType", () => {
    it("strips the #microsoft.graph. prefix", () => {
      expect(
        eventDetailType({ "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail" })
      ).toBe("chatRenamedEventMessageDetail");
    });

    it("tolerates a missing prefix or missing type", () => {
      expect(eventDetailType({ "@odata.type": "#foo" })).toBe("foo");
      expect(eventDetailType({})).toBe("unknown");
      expect(eventDetailType(undefined)).toBe("unknown");
    });
  });

  describe("isSystemEventMessage / isCallEventMessage", () => {
    it("detects system events by eventDetail even when messageType is unknownFutureValue", () => {
      const message = {
        ...callEnded("1", "2026-08-26T08:47:00Z", "PT1M"),
        messageType: "unknownFutureValue",
      } as ChatMessage;
      expect(isSystemEventMessage(message)).toBe(true);
      expect(isCallEventMessage(message)).toBe(true);
    });

    it("treats a plain message as neither", () => {
      const message = { id: "1", messageType: "message", body: { content: "hi" } } as ChatMessage;
      expect(isSystemEventMessage(message)).toBe(false);
      expect(isCallEventMessage(message)).toBe(false);
    });

    it("treats a non-call system event as system but not call", () => {
      const message = {
        id: "1",
        messageType: "systemEventMessage",
        eventDetail: { "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail" },
      } as unknown as ChatMessage;
      expect(isSystemEventMessage(message)).toBe(true);
      expect(isCallEventMessage(message)).toBe(false);
    });
  });

  describe("summarizeEventDetail", () => {
    it("maps callEnded fields and drops non-user participants", () => {
      const summary = summarizeEventDetail(
        callEnded("1", "2026-08-26T08:47:00Z", "PT39M29S").eventDetail
      );
      expect(summary).toEqual({
        type: "callEndedEventMessageDetail",
        callId: "call-1",
        callEventType: "call",
        callDuration: "PT39M29S",
        callDurationSeconds: 2369,
        initiator: { id: "user-a", displayName: undefined },
        callParticipants: [
          { id: "user-a", displayName: "Alice" },
          { id: "user-b", displayName: "Bob" },
        ],
      });
    });

    it("maps callStarted fields", () => {
      const summary = summarizeEventDetail(callStarted("1", "2026-08-26T08:07:00Z").eventDetail);
      expect(summary).toEqual({
        type: "callStartedEventMessageDetail",
        callId: "call-1",
        callEventType: "call",
        initiator: { id: "user-a", displayName: "Alice" },
      });
    });

    it("passes other event types through in raw", () => {
      const detail = {
        "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail",
        chatDisplayName: "New name",
      };
      expect(summarizeEventDetail(detail)).toEqual({
        type: "chatRenamedEventMessageDetail",
        raw: detail,
      });
    });

    it("returns undefined for missing detail", () => {
      expect(summarizeEventDetail(undefined)).toBeUndefined();
      expect(summarizeEventDetail(null)).toBeUndefined();
    });
  });

  describe("stitchCalls", () => {
    it("joins callStarted and callEnded on callId", () => {
      const calls = stitchCalls([
        callEnded("2", "2026-08-26T08:47:00Z", "PT39M29S"),
        callStarted("1", "2026-08-26T08:07:31Z"),
      ]);
      expect(calls).toEqual([
        {
          callId: "call-1",
          callEventType: "call",
          startDateTime: "2026-08-26T08:07:31Z",
          startEstimated: undefined,
          endDateTime: "2026-08-26T08:47:00Z",
          durationSeconds: 2369,
          initiator: { id: "user-a", displayName: "Alice" },
          participants: [
            { id: "user-a", displayName: "Alice" },
            { id: "user-b", displayName: "Bob" },
          ],
        },
      ]);
    });

    it("estimates start from end - duration when callStarted is missing", () => {
      const calls = stitchCalls([callEnded("2", "2026-08-26T08:47:00.000Z", "PT39M29S")]);
      expect(calls).toHaveLength(1);
      expect(calls[0].startEstimated).toBe(true);
      expect(calls[0].startDateTime).toBe("2026-08-26T08:07:31.000Z");
      expect(calls[0].endDateTime).toBe("2026-08-26T08:47:00.000Z");
      // initiator name is missing on callEnded but borrowed from the participant list
      expect(calls[0].initiator).toEqual({ id: "user-a", displayName: "Alice" });
    });

    it("treats an empty-string initiator displayName as absent", () => {
      const message = callEnded("2", "2026-08-26T08:47:00Z", "PT1M");
      (message.eventDetail as any).initiator.user.displayName = "";
      const summary = summarizeEventDetail(message.eventDetail);
      expect(summary?.initiator).toEqual({ id: "user-a", displayName: undefined });
    });

    it("keeps an open call (no callEnded) without end or duration", () => {
      const calls = stitchCalls([callStarted("1", "2026-08-26T12:22:00Z")]);
      expect(calls).toEqual([
        {
          callId: "call-1",
          callEventType: "call",
          startDateTime: "2026-08-26T12:22:00Z",
          startEstimated: undefined,
          initiator: { id: "user-a", displayName: "Alice" },
        },
      ]);
    });

    it("sorts calls newest first and ignores non-call events", () => {
      const calls = stitchCalls([
        callStarted("1", "2026-08-26T07:58:00Z", "call-old"),
        callEnded("2", "2026-08-26T08:47:00Z", "PT49M", "call-old"),
        {
          id: "3",
          messageType: "systemEventMessage",
          createdDateTime: "2026-08-26T09:00:00Z",
          eventDetail: { "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail" },
        } as unknown as ChatMessage,
        callStarted("4", "2026-08-26T12:22:00Z", "call-new", "meeting"),
        callEnded("5", "2026-08-26T12:30:00Z", "PT8M", "call-new", "meeting"),
        { id: "6", messageType: "message", body: { content: "hi" } } as ChatMessage,
      ]);
      expect(calls.map((c) => c.callId)).toEqual(["call-new", "call-old"]);
      expect(calls[0].callEventType).toBe("meeting");
      expect(calls[0].durationSeconds).toBe(480);
    });
  });
});
