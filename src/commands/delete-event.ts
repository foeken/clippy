import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getCalendarEvent, getCalendarEvents, deleteEvent, cancelEvent, type CalendarEvent } from '../lib/ews-client.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

function parseDay(day: string): Date {
  const now = new Date();

  switch (day.toLowerCase()) {
    case 'today':
      return now;
    case 'tomorrow':
      now.setDate(now.getDate() + 1);
      return now;
    case 'yesterday':
      now.setDate(now.getDate() - 1);
      return now;
    default:
      const parsed = new Date(day);
      return isNaN(parsed.getTime()) ? now : parsed;
  }
}

function writeError(message: string, json?: boolean): void {
  if (json) {
    console.log(JSON.stringify({ error: message }, null, 2));
  } else {
    console.error(`Error: ${message}`);
  }
}

export const deleteEventCommand = new Command('delete-event')
  .description('Delete/cancel a calendar event (sends cancellation if there are attendees)')
  .argument('[eventIndex]', 'Event index from the list (deprecated; use --id)')
  .option('--id <eventId>', 'Delete event by stable ID')
  .option('--day <day>', 'Day to show events from (today, tomorrow, YYYY-MM-DD)', 'today')
  .option('--search <text>', 'Search for events by title')
  .option('--message <text>', 'Cancellation message to send to attendees')
  .option('--force-delete', 'Delete without sending cancellation (even with attendees)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (eventIndex: string | undefined, options: {
    id?: string;
    day: string;
    search?: string;
    message?: string;
    forceDelete?: boolean;
    json?: boolean;
    token?: string;
  }) => {
    const authResult = await resolveAuth({
      token: options.token,
    });

    if (!authResult.success) {
      if (options.json) {
        console.log(JSON.stringify({ error: authResult.error }, null, 2));
      } else {
        console.error(`Error: ${authResult.error}`);
        console.error('\nCheck your .env file for EWS_CLIENT_ID and EWS_REFRESH_TOKEN.');
      }
      process.exit(1);
    }

    // If no id provided, list events
    if (!options.id) {
      // Get events for the day
      const baseDate = parseDay(options.day);
      const startOfDay = new Date(baseDate);
      startOfDay.setHours(0, 0, 0, 0);
      const endOfDay = new Date(baseDate);
      endOfDay.setHours(23, 59, 59, 999);

      const result = await getCalendarEvents(
        authResult.token!,
        startOfDay.toISOString(),
        endOfDay.toISOString()
      );

      if (!result.ok || !result.data) {
        writeError(result.error?.message || 'Failed to fetch events', options.json);
        process.exit(1);
      }

      // Filter to events the user owns (IsOrganizer) and optionally by search
      let events = result.data.filter(e => e.IsOrganizer && !e.IsCancelled);

      if (options.search) {
        const searchLower = options.search.toLowerCase();
        events = events.filter(e => e.Subject?.toLowerCase().includes(searchLower));
      }

      if (options.json) {
        console.log(JSON.stringify({
          events: events.map((e, i) => ({
            index: i + 1,
            id: e.Id,
            subject: e.Subject,
            start: e.Start.DateTime,
            end: e.End.DateTime,
          })),
        }, null, 2));
        return;
      }

      console.log(`\nYour events for ${formatDate(baseDate.toISOString())}:\n`);
      console.log('\u2500'.repeat(60));

      if (events.length === 0) {
        console.log('\n  No events found that you can delete.');
        console.log('  (You can only delete events you organized)\n');
        return;
      }

      for (let i = 0; i < events.length; i++) {
        const event = events[i];
        const startTime = formatTime(event.Start.DateTime);
        const endTime = formatTime(event.End.DateTime);
        const attendees = event.Attendees?.filter(a =>
          a.EmailAddress?.Address && a.Type !== 'Resource'
        ) || [];

        console.log(`\n  [${i + 1}] ${event.Subject}`);
        console.log(`      ${startTime} - ${endTime}`);
        console.log(`      ID: ${event.Id}`);
        if (event.Location?.DisplayName) {
          console.log(`      Location: ${event.Location.DisplayName}`);
        }
        if (attendees.length > 0) {
          console.log(`      Attendees: ${attendees.length} (will be notified on cancel)`);
        }
      }

      console.log('\n' + '\u2500'.repeat(60));
      console.log('\nTo delete/cancel an event:');
      console.log('  clippy delete-event <number>                    # Cancel & notify attendees');
      console.log('  clippy delete-event <number> --message "Sorry"  # With cancellation message');
      console.log('  clippy delete-event <number> --force-delete     # Delete without notifying');
      console.log('');
      return;
    }

    const targetResult = await getCalendarEvent(authResult.token!, options.id);
    if (!targetResult.ok || !targetResult.data) {
      writeError(targetResult.error?.message || `Event not found: ${options.id}`, options.json);
      process.exit(1);
    }

    const targetEvent: CalendarEvent = targetResult.data;
    if (targetEvent.IsCancelled) {
      writeError(`Cannot delete/cancel an already cancelled event: ${targetEvent.Subject}`, options.json);
      process.exit(1);
    }
    if (!targetEvent.IsOrganizer) {
      writeError('Cannot delete/cancel an event you do not organize. Use `clippy respond decline --id <eventId>` instead.', options.json);
      process.exit(1);
    }

    // Check if event has attendees (other than organizer)
    const attendees = targetEvent.Attendees?.filter(a =>
      a.EmailAddress?.Address && a.Type !== 'Resource'
    ) || [];
    const hasAttendees = attendees.length > 0;

    if (!options.json) {
      console.log(`\nDeleting: ${targetEvent.Subject}`);
      console.log(`  ${formatDate(targetEvent.Start.DateTime)} ${formatTime(targetEvent.Start.DateTime)} - ${formatTime(targetEvent.End.DateTime)}`);
    }

    let deleteResult;
    let action: string;

    if (hasAttendees && !options.forceDelete) {
      // Use cancel to send cancellation notices
      if (!options.json) {
        console.log(`  Attendees: ${attendees.map(a => a.EmailAddress?.Address).join(', ')}`);
        console.log(`  Sending cancellation notices...`);
      }
      deleteResult = await cancelEvent(authResult.token!, targetEvent.Id, options.message);
      action = 'cancelled';
    } else {
      // Just delete without notification
      deleteResult = await deleteEvent(authResult.token!, targetEvent.Id);
      action = 'deleted';
    }

    if (!deleteResult.ok) {
      if (options.json) {
        console.log(JSON.stringify({ error: deleteResult.error?.message || `Failed to ${action} event` }, null, 2));
      } else {
        console.error(`\nError: ${deleteResult.error?.message || `Failed to ${action} event`}`);
      }
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify({
        success: true,
        action,
        event: targetEvent.Subject,
        attendeesNotified: hasAttendees && !options.forceDelete ? attendees.length : 0,
      }, null, 2));
    } else {
      if (hasAttendees && !options.forceDelete) {
        console.log(`\n\u2713 Event cancelled. ${attendees.length} attendee(s) notified.\n`);
      } else {
        console.log('\n\u2713 Event deleted.\n');
      }
    }
  });
