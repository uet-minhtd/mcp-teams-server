import { describe, it, expect } from "vitest";
import {
  SearchUsersSchema,
  GetUserSchema,
  ListChannelsSchema,
  GetChannelMessagesSchema,
  SendChannelMessageSchema,
  GetChatMessagesSchema,
  SendChatMessageSchema,
  ListMeetingsSchema,
  GetMeetingSchema,
  GetMeetingAttendanceSchema,
  GetMeetingTranscriptsSchema,
  SearchMessagesSchema,
} from "../types.js";

describe("Types - Zod Schemas", () => {
  describe("SearchUsersSchema", () => {
    it("accepts valid input", () => {
      const result = SearchUsersSchema.safeParse({ query: "john" });
      expect(result.success).toBe(true);
    });

    it("rejects missing query", () => {
      const result = SearchUsersSchema.safeParse({});
      expect(result.success).toBe(false);
    });

    it("rejects non-string query", () => {
      const result = SearchUsersSchema.safeParse({ query: 123 });
      expect(result.success).toBe(false);
    });
  });

  describe("ListChannelsSchema", () => {
    it("accepts valid teamId", () => {
      const result = ListChannelsSchema.safeParse({ teamId: "team-123" });
      expect(result.success).toBe(true);
    });

    it("rejects missing teamId", () => {
      const result = ListChannelsSchema.safeParse({});
      expect(result.success).toBe(false);
    });
  });

  describe("GetChannelMessagesSchema", () => {
    it("accepts required fields", () => {
      const result = GetChannelMessagesSchema.safeParse({
        teamId: "t1",
        channelId: "c1",
      });
      expect(result.success).toBe(true);
    });

    it("applies default limit", () => {
      const result = GetChannelMessagesSchema.parse({
        teamId: "t1",
        channelId: "c1",
      });
      expect(result.limit).toBe(20);
    });

    it("accepts custom limit", () => {
      const result = GetChannelMessagesSchema.safeParse({
        teamId: "t1",
        channelId: "c1",
        limit: 50,
      });
      expect(result.success).toBe(true);
      if (result.success) expect(result.data.limit).toBe(50);
    });

    it("rejects missing teamId", () => {
      const result = GetChannelMessagesSchema.safeParse({
        channelId: "c1",
      });
      expect(result.success).toBe(false);
    });
  });

  describe("SendChannelMessageSchema", () => {
    it("accepts all required fields", () => {
      const result = SendChannelMessageSchema.safeParse({
        teamId: "t1",
        channelId: "c1",
        message: "Hello",
      });
      expect(result.success).toBe(true);
    });

    it("rejects missing message", () => {
      const result = SendChannelMessageSchema.safeParse({
        teamId: "t1",
        channelId: "c1",
      });
      expect(result.success).toBe(false);
    });
  });

  describe("GetChatMessagesSchema", () => {
    it("accepts valid chatId", () => {
      const result = GetChatMessagesSchema.safeParse({ chatId: "chat-1" });
      expect(result.success).toBe(true);
    });

    it("applies default limit", () => {
      const result = GetChatMessagesSchema.parse({ chatId: "chat-1" });
      expect(result.limit).toBe(20);
    });

    it("rejects missing chatId", () => {
      const result = GetChatMessagesSchema.safeParse({});
      expect(result.success).toBe(false);
    });
  });

  describe("SendChatMessageSchema", () => {
    it("accepts valid input", () => {
      const result = SendChatMessageSchema.safeParse({
        chatId: "chat-1",
        message: "Hi",
      });
      expect(result.success).toBe(true);
    });
  });

  describe("ListMeetingsSchema", () => {
    it("accepts empty input with default limit", () => {
      const result = ListMeetingsSchema.parse({});
      expect(result.limit).toBe(20);
    });

    it("accepts custom limit", () => {
      const result = ListMeetingsSchema.parse({ limit: 10 });
      expect(result.limit).toBe(10);
    });
  });

  describe("GetMeetingSchema", () => {
    it("accepts valid meetingId", () => {
      const result = GetMeetingSchema.safeParse({ meetingId: "meeting-1" });
      expect(result.success).toBe(true);
    });

    it("rejects missing meetingId", () => {
      const result = GetMeetingSchema.safeParse({});
      expect(result.success).toBe(false);
    });
  });

  describe("GetMeetingAttendanceSchema", () => {
    it("accepts valid meetingId", () => {
      const result = GetMeetingAttendanceSchema.safeParse({ meetingId: "m1" });
      expect(result.success).toBe(true);
    });
  });

  describe("GetMeetingTranscriptsSchema", () => {
    it("accepts valid meetingId", () => {
      const result = GetMeetingTranscriptsSchema.safeParse({ meetingId: "m1" });
      expect(result.success).toBe(true);
    });
  });

  describe("SearchMessagesSchema", () => {
    it("accepts query with default limit", () => {
      const result = SearchMessagesSchema.parse({ query: "hello" });
      expect(result.query).toBe("hello");
      expect(result.limit).toBe(20);
    });

    it("rejects missing query", () => {
      const result = SearchMessagesSchema.safeParse({});
      expect(result.success).toBe(false);
    });
  });

  describe("GetUserSchema", () => {
    it("accepts valid userId", () => {
      const result = GetUserSchema.safeParse({ userId: "user-1" });
      expect(result.success).toBe(true);
    });

    it("rejects missing userId", () => {
      const result = GetUserSchema.safeParse({});
      expect(result.success).toBe(false);
    });
  });
});
