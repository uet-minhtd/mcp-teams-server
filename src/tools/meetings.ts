import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../graph/client.js";
import {
  ListMeetingsSchema,
  GetMeetingSchema,
  GetMeetingAttendanceSchema,
  GetMeetingTranscriptsSchema,
} from "../types.js";

export function registerMeetingsTools(
  server: McpServer,
  graphService: GraphService,
  userToken?: string
): void {
  server.registerTool(
    "list_meetings",
    {
      description: "List the user's calendar events including online meetings (upcoming and past). Read-only, no data is modified.",
      inputSchema: ListMeetingsSchema.shape,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ limit }) => {
      try {
        const client = await graphService.getClient(userToken);
        // /me/onlineMeetings does not support $top or listing without $filter.
        // Use /me/events filtered by isOnlineMeeting instead — supports full OData query params.
        const result = await client
          .api("/me/events")
          .orderby("start/dateTime desc")
          .top(limit)
          .select("id,subject,start,end,isOnlineMeeting,onlineMeeting,organizer,attendees,webLink")
          .get();
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            { type: "text" as const, text: `Failed to list meetings: ${message}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.registerTool(
    "get_meeting",
    {
      description: "Get details of a specific meeting or calendar event by ID. Read-only, no data is modified.",
      inputSchema: GetMeetingSchema.shape,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ meetingId }) => {
      try {
        const client = await graphService.getClient(userToken);
        // meetingId can be either a calendar event ID or an onlineMeeting ID.
        // Try /me/events first (calendar-based meetings), then fall back to /me/onlineMeetings.
        let result: unknown;
        try {
          result = await client
            .api(`/me/events/${meetingId}`)
            .select("id,subject,start,end,isOnlineMeeting,onlineMeeting,organizer,attendees,webLink,bodyPreview")
            .get();
        } catch {
          // Fall back to standalone onlineMeeting (created via Graph API directly)
          result = await client
            .api(`/me/onlineMeetings/${meetingId}`)
            .get();
        }
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            { type: "text" as const, text: `Failed to get meeting: ${message}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.registerTool(
    "get_meeting_attendance",
    {
      description: "Get attendance reports for a Teams meeting. Read-only. " +
        "⚠️ Contains sensitive participant data (join/leave times, duration). Only call when the user explicitly requests attendance information.",
      inputSchema: GetMeetingAttendanceSchema.shape,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ meetingId }) => {
      try {
        const client = await graphService.getClient(userToken);
        const reports = await client
          .api(`/me/onlineMeetings/${meetingId}/attendanceReports`)
          .get();

        const reportsList = reports.value as Array<{ id: string }>;
        const detailed = [];

        for (const report of reportsList) {
          const records = await client
            .api(
              `/me/onlineMeetings/${meetingId}/attendanceReports/${report.id}/attendanceRecords`
            )
            .get();
          detailed.push({
            reportId: report.id,
            records: records.value,
          });
        }

        return {
          content: [{ type: "text" as const, text: JSON.stringify(detailed, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            {
              type: "text" as const,
              text: `Failed to get attendance: ${message}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  server.registerTool(
    "get_meeting_transcripts",
    {
      description: "Get full transcripts of a Teams meeting (requires Teams Premium recording). Read-only. " +
        "⚠️ Contains sensitive conversation content. Only call when the user explicitly requests transcript data.",
      inputSchema: GetMeetingTranscriptsSchema.shape,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ meetingId }) => {
      try {
        const client = await graphService.getClient(userToken);
        const result = await client
          .api(`/me/onlineMeetings/${meetingId}/transcripts`)
          .get();

        const transcripts = result.value as Array<{
          id: string;
          createdDateTime: string;
        }>;
        const detailed = [];

        for (const transcript of transcripts) {
          try {
            const content = await client
              .api(
                `/me/onlineMeetings/${meetingId}/transcripts/${transcript.id}/content`
              )
              .get();
            detailed.push({
              transcriptId: transcript.id,
              createdDateTime: transcript.createdDateTime,
              content: JSON.stringify(content),
            });
          } catch {
            detailed.push({
              transcriptId: transcript.id,
              createdDateTime: transcript.createdDateTime,
              content: "Unable to fetch transcript content",
            });
          }
        }

        return {
          content: [{ type: "text" as const, text: JSON.stringify(detailed, null, 2) }],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [
            {
              type: "text" as const,
              text: `Failed to get transcripts: ${message}`,
            },
          ],
          isError: true,
        };
      }
    }
  );
}
