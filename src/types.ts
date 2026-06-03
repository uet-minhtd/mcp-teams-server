import { z } from "zod";

// Tool input schemas

export const SearchUsersSchema = z.object({
  query: z.string().describe("Name or email to search for"),
});

export const GetUserSchema = z.object({
  userId: z.string().describe("User ID or userPrincipalName"),
});

export const ListChannelsSchema = z.object({
  teamId: z.string().describe("Team ID"),
});

export const GetChannelMessagesSchema = z.object({
  teamId: z.string().describe("Team ID"),
  channelId: z.string().describe("Channel ID"),
  limit: z.number().optional().default(20).describe("Max messages to return"),
});

export const SendChannelMessageSchema = z.object({
  teamId: z.string().describe("Team ID"),
  channelId: z.string().describe("Channel ID"),
  message: z.string().describe("Message content"),
});

export const GetChatMessagesSchema = z.object({
  chatId: z.string().describe("Chat ID"),
  limit: z.number().optional().default(20).describe("Max messages to return"),
});

export const SendChatMessageSchema = z.object({
  chatId: z.string().describe("Chat ID"),
  message: z.string().describe("Message content"),
});

export const ListMeetingsSchema = z.object({
  limit: z.number().optional().default(20).describe("Max meetings to return"),
});

export const GetMeetingSchema = z.object({
  meetingId: z.string().describe("Meeting ID"),
});

export const GetMeetingAttendanceSchema = z.object({
  meetingId: z.string().describe("Meeting ID"),
});

export const GetMeetingTranscriptsSchema = z.object({
  meetingId: z.string().describe("Meeting ID"),
});

export const SearchMessagesSchema = z.object({
  query: z.string().describe("Search query (KQL syntax)"),
  limit: z.number().optional().default(20).describe("Max results to return"),
});

// Auth context attached to request by middleware

export interface AuthContext {
  token: string;
  userId: string;
  displayName: string;
  userPrincipalName: string;
}
