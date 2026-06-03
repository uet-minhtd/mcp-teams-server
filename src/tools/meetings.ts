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
  server.tool(
    "list_meetings",
    "List online meetings (upcoming and past)",
    ListMeetingsSchema.shape,
    async ({ limit }) => {
      try {
        const client = await graphService.getClient(userToken);
        const result = await client
          .api("/me/onlineMeetings")
          .top(limit)
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

  server.tool(
    "get_meeting",
    "Get details of a specific online meeting",
    GetMeetingSchema.shape,
    async ({ meetingId }) => {
      try {
        const client = await graphService.getClient(userToken);
        const result = await client
          .api(`/me/onlineMeetings/${meetingId}`)
          .get();
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

  server.tool(
    "get_meeting_attendance",
    "Get attendance reports and records for a meeting",
    GetMeetingAttendanceSchema.shape,
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

  server.tool(
    "get_meeting_transcripts",
    "Get transcripts for a meeting (requires Teams Premium recording)",
    GetMeetingTranscriptsSchema.shape,
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
