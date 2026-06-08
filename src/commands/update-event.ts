import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getCalendarEvents, updateEvent, searchRooms, getRooms, getCalendarEvent, type CalendarShowAs } from '../lib/ews-client.js';

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

function parseTimeToDate(timeStr: string, baseDate: Date): Date {
  const result = new Date(baseDate);

  const timeMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  if (timeMatch) {
    result.setHours(parseInt(timeMatch[1]), parseInt(timeMatch[2]), 0, 0);
    return result;
  }

  const hourMatch = timeStr.match(/^(\d{1,2})(am|pm)?$/i);
  if (hourMatch) {
    let hour = parseInt(hourMatch[1]);
    const isPM = hourMatch[2]?.toLowerCase() === 'pm';
    if (isPM && hour < 12) hour += 12;
    if (!isPM && hour === 12) hour = 0;
    result.setHours(hour, 0, 0, 0);
    return result;
  }

  return result;
}

function toEwsDateTime(date: Date): string {
  // EWS can preserve an item's existing timezone when given floating timestamps.
  // The CLI accepts local wall-clock input, so send the exact UTC instant.
  return date.toISOString();
}

function formatShowAs(showAs?: string): string {
  switch (showAs) {
    case 'OOF':
      return 'Out of office';
    case 'WorkingElsewhere':
      return 'Working elsewhere';
    default:
      return showAs || 'Busy';
  }
}

function formatReminder(event: { reminderIsSet?: boolean; reminderMinutesBeforeStart?: number }): string {
  if (event.reminderIsSet === false) return 'off';
  if (event.reminderIsSet === true) {
    return event.reminderMinutesBeforeStart !== undefined
      ? `${event.reminderMinutesBeforeStart} minutes before`
      : 'on';
  }
  return 'unknown';
}

function parseShowAs(value: string): CalendarShowAs | undefined {
  const normalized = value.trim().toLowerCase().replace(/[\s_]+/g, '-');
  const aliases: Record<string, CalendarShowAs> = {
    free: 'Free',
    busy: 'Busy',
    tentative: 'Tentative',
    oof: 'OOF',
    away: 'OOF',
    'out-of-office': 'OOF',
    workingelsewhere: 'WorkingElsewhere',
    'working-elsewhere': 'WorkingElsewhere',
    elsewhere: 'WorkingElsewhere',
  };

  return aliases[normalized];
}

function resolveShowAs(options: { showAs?: string; free?: boolean; busy?: boolean }): CalendarShowAs | undefined {
  const values: CalendarShowAs[] = [];

  if (options.showAs) {
    const showAs = parseShowAs(options.showAs);
    if (!showAs) {
      throw new Error('Invalid --show-as value. Use free, busy, tentative, oof, or working-elsewhere.');
    }
    values.push(showAs);
  }
  if (options.free) values.push('Free');
  if (options.busy) values.push('Busy');

  const unique = [...new Set(values)];
  if (unique.length > 1) {
    throw new Error('Conflicting availability options. Use only one of --show-as, --free, or --busy.');
  }

  return unique[0];
}

function resolveReminder(options: { reminder?: string | false }): {
  hasReminderUpdate: boolean;
  reminderIsSet?: boolean;
  reminderMinutesBeforeStart?: number;
} {
  if (options.reminder === undefined) {
    return { hasReminderUpdate: false };
  }

  if (options.reminder === false) {
    return { hasReminderUpdate: true, reminderIsSet: false };
  }

  const value = options.reminder.trim();
  if (!/^\d+$/.test(value)) {
    throw new Error('Invalid --reminder value. Use a non-negative number of minutes.');
  }

  return {
    hasReminderUpdate: true,
    reminderIsSet: true,
    reminderMinutesBeforeStart: Number.parseInt(value, 10),
  };
}

function writeError(message: string, json?: boolean): void {
  if (json) {
    console.log(JSON.stringify({ error: message }, null, 2));
  } else {
    console.error(`Error: ${message}`);
  }
}

export const updateEventCommand = new Command('update-event')
  .description('Update a calendar event')
  .argument('[eventIndex]', 'Event index from the list (deprecated; use --id)')
  .option('--id <eventId>', 'Update event by stable ID')
  .option('--day <day>', 'Day to show events from (today, tomorrow, YYYY-MM-DD)', 'today')
  .option('--title <text>', 'New title/subject')
  .option('--description <text>', 'New description/body')
  .option('--start <time>', 'New start time (e.g., 14:00, 2pm)')
  .option('--end <time>', 'New end time (e.g., 15:00, 3pm)')
  .option('--add-attendee <email>', 'Add an attendee (can be used multiple times)', (val, arr: string[]) => [...arr, val], [])
  .option('--room <room>', 'Set/change meeting room (name or email)')
  .option('--location <text>', 'Set location text')
  .option('--show-as <status>', 'Set availability: free, busy, tentative, oof, working-elsewhere')
  .option('--free', 'Shortcut for --show-as free')
  .option('--busy', 'Shortcut for --show-as busy')
  .option('--reminder <minutes>', 'Set a reminder this many minutes before the event')
  .option('--no-reminder', 'Clear/disable the event reminder')
  .option('--teams', 'Make it a Teams meeting')
  .option('--no-teams', 'Remove Teams meeting')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (eventIndex: string | undefined, options: {
    id?: string;
    day: string;
    title?: string;
    description?: string;
    start?: string;
    end?: string;
    addAttendee: string[];
    room?: string;
    location?: string;
    showAs?: string;
    free?: boolean;
    busy?: boolean;
    reminder?: string | false;
    teams?: boolean;
    json?: boolean;
    token?: string;
  }) => {
    let requestedShowAs: CalendarShowAs | undefined;
    let requestedReminder: ReturnType<typeof resolveReminder>;
    try {
      requestedShowAs = resolveShowAs(options);
      requestedReminder = resolveReminder(options);
    } catch (err) {
      writeError(err instanceof Error ? err.message : 'Invalid update option', options.json);
      process.exit(1);
    }

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
      if (options.json) {
        console.log(JSON.stringify({ error: result.error?.message || 'Failed to fetch events' }, null, 2));
      } else {
        console.error(`Error: ${result.error?.message || 'Failed to fetch events'}`);
      }
      process.exit(1);
    }

    // Include attendee-owned copies so reminder-only updates can target them.
    const events = result.data.filter(e => !e.IsCancelled);

    // If no target provided, list events
    if (!options.id && !eventIndex) {
      if (options.json) {
        console.log(JSON.stringify({
          events: events.map((e, i) => ({
            index: i + 1,
            id: e.Id,
            subject: e.Subject,
            start: e.Start.DateTime,
            end: e.End.DateTime,
            showAs: e.ShowAs,
            isOrganizer: e.IsOrganizer,
            reminderIsSet: e.reminderIsSet,
            reminderMinutesBeforeStart: e.reminderMinutesBeforeStart,
            attendees: e.Attendees?.map(a => a.EmailAddress?.Address),
          })),
        }, null, 2));
        return;
      }

      console.log(`\nYour events for ${formatDate(baseDate.toISOString())}:\n`);
      console.log('\u2500'.repeat(60));

      if (events.length === 0) {
        console.log('\n  No events found.\n');
        return;
      }

      for (let i = 0; i < events.length; i++) {
        const event = events[i];
        const startTime = formatTime(event.Start.DateTime);
        const endTime = formatTime(event.End.DateTime);

        console.log(`\n  [${i + 1}] ${event.Subject}`);
        console.log(`      ${startTime} - ${endTime}`);
        console.log(`      Show as: ${formatShowAs(event.ShowAs)}`);
        console.log(`      Reminder: ${formatReminder(event)}`);
        if (!event.IsOrganizer) {
          console.log('      Meeting edits: organizer only; reminder updates allowed');
        }
        console.log(`      ID: ${event.Id}`);
        if (event.Location?.DisplayName) {
          console.log(`      Location: ${event.Location.DisplayName}`);
        }
        if (event.Attendees && event.Attendees.length > 0) {
          const attendeeList = event.Attendees
            .filter(a => a.Type !== 'Resource')
            .map(a => a.EmailAddress?.Address)
            .filter(Boolean);
          if (attendeeList.length > 0) {
            console.log(`      Attendees: ${attendeeList.join(', ')}`);
          }
        }
      }

      console.log('\n' + '\u2500'.repeat(60));
      console.log('\nTo update an event:');
      console.log('  clippy update-event <number> --title "New Title"');
      console.log('  clippy update-event <number> --add-attendee user@example.com');
      console.log('  clippy update-event <number> --room "Taxi"');
      console.log('  clippy update-event <number> --start 14:00 --end 15:00');
      console.log('  clippy update-event <number> --show-as busy');
      console.log('  clippy update-event <number> --reminder 30');
      console.log('  clippy update-event <number> --no-reminder');
      console.log('');
      return;
    }

    // Get the target event by ID or list index
    let targetEvent = options.id ? events.find(e => e.Id === options.id) : undefined;
    if (options.id && !targetEvent) {
      const targetResult = await getCalendarEvent(authResult.token!, options.id);
      if (!targetResult.ok || !targetResult.data) {
        writeError(`Invalid event id: ${options.id}`, options.json);
        process.exit(1);
      }
      targetEvent = targetResult.data;
      if (targetEvent.IsCancelled) {
        writeError('Cannot update a cancelled event.', options.json);
        process.exit(1);
      }
    }
    if (!options.id) {
      const index = Number.parseInt(eventIndex || '', 10) - 1;
      if (!Number.isInteger(index) || index < 0 || index >= events.length) {
        writeError(`Invalid event index: ${eventIndex}`, options.json);
        process.exit(1);
      }
      targetEvent = events[index];
    }

    if (!targetEvent) {
      writeError('No event selected.', options.json);
      process.exit(1);
    }

    // Check if any update options were provided
    const hasOrganizerUpdates = options.title || options.description || options.start ||
      options.end || options.addAttendee.length > 0 || options.room ||
      options.location || requestedShowAs !== undefined || options.teams !== undefined;
    const hasUpdates = hasOrganizerUpdates || requestedReminder.hasReminderUpdate;

    if (!hasUpdates) {
      // Show current event details
      console.log(`\nEvent: ${targetEvent.Subject}`);
      console.log(`  When: ${formatDate(targetEvent.Start.DateTime)} ${formatTime(targetEvent.Start.DateTime)} - ${formatTime(targetEvent.End.DateTime)}`);
      console.log(`  Show as: ${formatShowAs(targetEvent.ShowAs)}`);
      console.log(`  Reminder: ${formatReminder(targetEvent)}`);
      if (targetEvent.Location?.DisplayName) {
        console.log(`  Location: ${targetEvent.Location.DisplayName}`);
      }
      if (targetEvent.Attendees && targetEvent.Attendees.length > 0) {
        console.log('  Attendees:');
        for (const a of targetEvent.Attendees) {
          const typeLabel = a.Type === 'Resource' ? ' (Room)' : '';
          console.log(`    - ${a.EmailAddress?.Address}${typeLabel}`);
        }
      }
      console.log('\nUse options like --title, --add-attendee, --room, --show-as, or --reminder to update.');
      return;
    }

    if (hasOrganizerUpdates && !targetEvent.IsOrganizer) {
      writeError('Cannot update organizer-controlled fields on an event you did not organize. Reminder-only updates are allowed.', options.json);
      process.exit(1);
    }

    // Build update payload
    const updateOptions: Parameters<typeof updateEvent>[0] = {
      token: authResult.token!,
      eventId: targetEvent.Id,
    };

    if (options.title) {
      updateOptions.subject = options.title;
    }

    if (options.description) {
      updateOptions.body = options.description;
    }

    // Handle time changes
    if (options.start || options.end) {
      const eventDate = new Date(targetEvent.Start.DateTime);

      if (options.start) {
        const newStart = parseTimeToDate(options.start, eventDate);
        updateOptions.start = toEwsDateTime(newStart);
      }

      if (options.end) {
        const newEnd = parseTimeToDate(options.end, eventDate);
        updateOptions.end = toEwsDateTime(newEnd);
      }
    }

    // Handle location
    if (options.location) {
      updateOptions.location = options.location;
    }

    if (requestedShowAs) {
      updateOptions.showAs = requestedShowAs;
    }

    if (requestedReminder.hasReminderUpdate) {
      updateOptions.reminderIsSet = requestedReminder.reminderIsSet;
      updateOptions.reminderMinutesBeforeStart = requestedReminder.reminderMinutesBeforeStart;
    }

    // Handle room
    let roomEmail: string | undefined;
    let roomName: string | undefined;

    if (options.room) {
      if (options.room.includes('@')) {
        roomEmail = options.room;
        roomName = options.room;
      } else {
        let roomsResult = await searchRooms(authResult.token!, options.room);
        if (!roomsResult.ok || !roomsResult.data || roomsResult.data.length === 0) {
          roomsResult = await getRooms(authResult.token!);
        }

        if (roomsResult.ok && roomsResult.data) {
          const found = roomsResult.data.find(
            r => r.Name.toLowerCase().includes(options.room!.toLowerCase())
          );
          if (found) {
            roomEmail = found.Address;
            roomName = found.Name;
          } else {
            console.error(`Room not found: ${options.room}`);
            process.exit(1);
          }
        }
      }

      if (roomName) {
        updateOptions.location = roomName;
      }
    }

    // Handle attendees (merge existing with new)
    if (options.addAttendee.length > 0 || roomEmail) {
      const existingAttendees: Array<{ email: string; name?: string; type: 'Required' | 'Optional' | 'Resource' }> = (targetEvent.Attendees || []).map(a => ({
        email: a.EmailAddress?.Address || '',
        name: a.EmailAddress?.Name,
        type: a.Type as 'Required' | 'Optional' | 'Resource',
      }));

      // Add new attendees
      for (const email of options.addAttendee) {
        if (!existingAttendees.find(a => a.email.toLowerCase() === email.toLowerCase())) {
          existingAttendees.push({ email, type: 'Required' });
        }
      }

      // Add room if specified
      if (roomEmail) {
        // Remove any existing room
        const withoutRooms = existingAttendees.filter(a => a.type !== 'Resource');
        withoutRooms.push({ email: roomEmail, name: roomName, type: 'Resource' });
        updateOptions.attendees = withoutRooms;
      } else {
        updateOptions.attendees = existingAttendees;
      }
    }

    // Handle Teams
    if (options.teams !== undefined) {
      updateOptions.isOnlineMeeting = options.teams;
    }

    console.log(`\nUpdating: ${targetEvent.Subject}`);

    const updateResult = await updateEvent(updateOptions);

    if (!updateResult.ok) {
      if (options.json) {
        console.log(JSON.stringify({ error: updateResult.error?.message || 'Failed to update event' }, null, 2));
      } else {
        console.error(`\nError: ${updateResult.error?.message || 'Failed to update event'}`);
      }
      process.exit(1);
    }

    if (options.json) {
      console.log(JSON.stringify({
        success: true,
        event: {
          id: updateResult.data?.Id || targetEvent.Id,
          subject: updateResult.data?.Subject || targetEvent.Subject,
          start: updateResult.data?.Start.DateTime || targetEvent.Start.DateTime,
          end: updateResult.data?.End.DateTime || targetEvent.End.DateTime,
          showAs: updateResult.data?.ShowAs || requestedShowAs || targetEvent.ShowAs,
          reminderIsSet: updateResult.data?.reminderIsSet ?? (requestedReminder.hasReminderUpdate ? requestedReminder.reminderIsSet : targetEvent.reminderIsSet),
          reminderMinutesBeforeStart: updateResult.data?.reminderMinutesBeforeStart ?? (requestedReminder.hasReminderUpdate ? requestedReminder.reminderMinutesBeforeStart : targetEvent.reminderMinutesBeforeStart),
        },
      }, null, 2));
    } else {
      console.log('\n\u2713 Event updated successfully.\n');
      const resultSubject = updateResult.data?.Subject || targetEvent.Subject;
      const resultStart = updateResult.data?.Start.DateTime || targetEvent.Start.DateTime;
      const resultEnd = updateResult.data?.End.DateTime || targetEvent.End.DateTime;
      const resultShowAs = updateResult.data?.ShowAs || requestedShowAs || targetEvent.ShowAs;
      const resultReminderIsSet = updateResult.data?.reminderIsSet ?? (requestedReminder.hasReminderUpdate ? requestedReminder.reminderIsSet : targetEvent.reminderIsSet);
      const resultReminderMinutes = updateResult.data?.reminderMinutesBeforeStart ?? (requestedReminder.hasReminderUpdate ? requestedReminder.reminderMinutesBeforeStart : targetEvent.reminderMinutesBeforeStart);

      console.log(`  Title: ${resultSubject}`);
      console.log(`  When:  ${formatDate(resultStart)} ${formatTime(resultStart)} - ${formatTime(resultEnd)}`);
      console.log(`  Show as: ${formatShowAs(resultShowAs)}`);
      console.log(`  Reminder: ${formatReminder({ reminderIsSet: resultReminderIsSet, reminderMinutesBeforeStart: resultReminderMinutes })}`);
      console.log('');
    }
  });
