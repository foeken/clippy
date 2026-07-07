import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getCalendarEvent, getCalendarEvents, respondToEvent, getOwaUserInfo, ResponseType, type CalendarAttendee, type CalendarEvent } from '../lib/ews-client.js';
import { assertReadWriteAllowed } from '../lib/readonly.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

function parseTimeToDate(timeStr: string, baseDate: Date): Date {
  const result = new Date(baseDate);

  const timeMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  if (timeMatch) {
    result.setHours(Number.parseInt(timeMatch[1], 10), Number.parseInt(timeMatch[2], 10), 0, 0);
    return result;
  }

  const hourMatch = timeStr.match(/^(\d{1,2})(am|pm)?$/i);
  if (hourMatch) {
    let hour = Number.parseInt(hourMatch[1], 10);
    const isPM = hourMatch[2]?.toLowerCase() === 'pm';
    if (isPM && hour < 12) hour += 12;
    if (!isPM && hour === 12) hour = 0;
    result.setHours(hour, 0, 0, 0);
    return result;
  }

  const parsed = new Date(timeStr);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  throw new Error('Invalid time value. Use HH:MM, HHam/pm, YYYY-MM-DDTHH:MM, or an ISO timestamp.');
}

function parseDateTimeToDate(value: string, baseDate: Date): Date {
  const trimmed = value.trim();
  const localDateTime = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s])(\d{1,2}):(\d{2})$/);
  if (localDateTime) {
    const [, year, month, day, hour, minute] = localDateTime;
    return new Date(
      Number.parseInt(year, 10),
      Number.parseInt(month, 10) - 1,
      Number.parseInt(day, 10),
      Number.parseInt(hour, 10),
      Number.parseInt(minute, 10),
      0,
      0
    );
  }

  if (/^\d{4}-\d{2}-\d{2}T/.test(trimmed)) {
    const parsed = new Date(trimmed);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }

  return parseTimeToDate(trimmed, baseDate);
}

function parseProposalDate(value: string): Date {
  const normalized = value.trim().toLowerCase();
  const result = new Date();

  if (normalized === 'today') {
    result.setHours(0, 0, 0, 0);
    return result;
  }
  if (normalized === 'tomorrow') {
    result.setDate(result.getDate() + 1);
    result.setHours(0, 0, 0, 0);
    return result;
  }
  if (normalized === 'yesterday') {
    result.setDate(result.getDate() - 1);
    result.setHours(0, 0, 0, 0);
    return result;
  }

  const dateOnly = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!dateOnly) {
    throw new Error('Invalid --date value. Use YYYY-MM-DD, today, tomorrow, or yesterday.');
  }

  const [, year, month, day] = dateOnly;
  return new Date(
    Number.parseInt(year, 10),
    Number.parseInt(month, 10) - 1,
    Number.parseInt(day, 10),
    0,
    0,
    0,
    0
  );
}

function moveDateKeepingTime(date: Date, targetDate: Date): Date {
  return new Date(
    targetDate.getFullYear(),
    targetDate.getMonth(),
    targetDate.getDate(),
    date.getHours(),
    date.getMinutes(),
    date.getSeconds(),
    date.getMilliseconds()
  );
}

export function resolveProposedResponseWindow(
  event: CalendarEvent,
  options: { date?: string; start?: string; end?: string; duration?: string }
): { start: Date; end: Date } {
  const currentStart = new Date(event.Start.DateTime);
  const currentEnd = new Date(event.End.DateTime);
  if (Number.isNaN(currentStart.getTime()) || Number.isNaN(currentEnd.getTime())) {
    throw new Error('Event has invalid start or end time.');
  }

  const durationMs = currentEnd.getTime() - currentStart.getTime();
  if (durationMs <= 0) {
    throw new Error('Event has invalid duration.');
  }

  const proposalDate = options.date ? parseProposalDate(options.date) : undefined;
  const baseDate = proposalDate || currentStart;
  let start = proposalDate ? moveDateKeepingTime(currentStart, proposalDate) : new Date(currentStart);
  let end = proposalDate ? new Date(start.getTime() + durationMs) : new Date(currentEnd);

  if (options.start) {
    start = parseDateTimeToDate(options.start, baseDate);
    if (!options.end) {
      end = new Date(start.getTime() + durationMs);
    }
  }

  if (options.end) {
    end = parseDateTimeToDate(options.end, start);
  }

  if (options.duration !== undefined) {
    if (options.end) {
      throw new Error('Use only one of --end or --duration for a proposal.');
    }
    const durationMinutes = Number.parseInt(options.duration, 10);
    if (!/^\d+$/.test(options.duration.trim()) || durationMinutes <= 0) {
      throw new Error('Invalid --duration value. Use a positive number of minutes.');
    }
    end = new Date(start.getTime() + durationMinutes * 60 * 1000);
  }

  if (end <= start) {
    throw new Error('Proposed event end must be after the proposed start.');
  }

  return { start, end };
}

function getResponseIcon(response: string): string {
  switch (response) {
    case 'Accepted': return '\u2713';
    case 'Declined': return '\u2717';
    case 'TentativelyAccepted': return '?';
    case 'None':
    case 'NotResponded': return '\u2022';
    default: return ' ';
  }
}

function writeError(message: string, json?: boolean): void {
  if (json) {
    console.log(JSON.stringify({ error: message }, null, 2));
  } else {
    console.error(`Error: ${message}`);
  }
}

function getMyAttendance(event: CalendarEvent, userEmail?: string): CalendarAttendee | undefined {
  if (!userEmail) return undefined;
  return event.Attendees?.find(
    a => a.EmailAddress?.Address?.toLowerCase() === userEmail
  );
}

export function normalizeMyResponseType(responseType?: string): CalendarAttendee['Status']['Response'] | undefined {
  switch (responseType) {
    case 'Accept':
    case 'Accepted':
      return 'Accepted';
    case 'Decline':
    case 'Declined':
      return 'Declined';
    case 'Tentative':
    case 'TentativelyAccepted':
      return 'TentativelyAccepted';
    case 'NoResponseReceived':
    case 'NotResponded':
      return 'NotResponded';
    case 'Organizer':
      return 'Organizer';
    case 'Unknown':
    case 'None':
      return 'None';
    default:
      return undefined;
  }
}

export function getInvitationResponseStatus(event: CalendarEvent, userEmail?: string): CalendarAttendee['Status']['Response'] | undefined {
  return normalizeMyResponseType(event.MyResponseType) || getMyAttendance(event, userEmail)?.Status?.Response;
}

export function isCancellationNoticeSubject(subject?: string): boolean {
  const normalized = subject?.trim().toLowerCase() || '';
  return /^(canceled|cancelled|geannuleerd)\s*:/.test(normalized);
}

export function isPendingInvitation(event: CalendarEvent, userEmail?: string, onlyRequired = false): boolean {
  if (event.IsCancelled) return false;
  if (isCancellationNoticeSubject(event.Subject)) return false;
  if (event.IsOrganizer) return false;

  const response = getInvitationResponseStatus(event, userEmail);
  const isPending = response === 'None' || response === 'NotResponded';
  if (!isPending) return false;

  const myAttendance = getMyAttendance(event, userEmail);
  const isOptional = myAttendance?.Type === 'Optional';
  if (onlyRequired && isOptional) return false;

  return true;
}

export const respondCommand = new Command('respond')
  .description('Respond to calendar invitations (accept/decline/tentative/propose)')
  .argument('[action]', 'Action: list, accept, decline, tentative, propose')
  .argument('[eventIndex]', 'Event index from the list (deprecated; use --id)')
  .option('--id <eventId>', 'Respond to a specific event by stable ID')
  .option('--comment <text>', 'Add a comment to your response')
  .option('--date <date>', 'For propose: proposed date (YYYY-MM-DD, today, tomorrow, yesterday)')
  .option('--start <time>', 'For propose: proposed start time (HH:MM, YYYY-MM-DDTHH:MM, or ISO timestamp)')
  .option('--end <time>', 'For propose: proposed end time (HH:MM, YYYY-MM-DDTHH:MM, or ISO timestamp)')
  .option('--duration <minutes>', 'For propose: proposed duration in minutes')
  .option('--no-notify', 'Don\'t send response to organizer')
  .option('--include-optional', 'Include optional invitations (default)', true)
  .option('--only-required', 'Only show required invitations')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (action: string | undefined, eventIndex: string | undefined, options: {
    id?: string;
    comment?: string;
    date?: string;
    start?: string;
    end?: string;
    duration?: string;
    notify: boolean;
    includeOptional?: boolean;
    onlyRequired?: boolean;
    json?: boolean;
    token?: string;
  }) => {
    // Default action is 'list'
    const actionLower = (action || 'list').toLowerCase();

    if (!['list', 'accept', 'decline', 'tentative', 'propose'].includes(actionLower)) {
      writeError(`Unknown action: ${action}`, options.json);
      if (!options.json) {
        console.error('Valid actions: list, accept, decline, tentative, propose');
      }
      process.exit(1);
    }

    if (actionLower !== 'list') {
      assertReadWriteAllowed('Responding to invitations');
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

    if (actionLower === 'list') {
      // Get user's email to identify their response status
      const userInfo = await getOwaUserInfo(authResult.token!);
      const userEmail = userInfo.data?.email?.toLowerCase();

      // Fetch upcoming events
      const now = new Date();
      const futureDate = new Date(now);
      futureDate.setDate(futureDate.getDate() + 30); // Look 30 days ahead

      const result = await getCalendarEvents(
        authResult.token!,
        now.toISOString(),
        futureDate.toISOString()
      );

      if (!result.ok || !result.data) {
        writeError(result.error?.message || 'Failed to fetch events', options.json);
        process.exit(1);
      }

      // Filter to events where user is an attendee (and not organizer)
      const pendingEvents = result.data.filter(event => {
        return isPendingInvitation(event, userEmail, Boolean(options.onlyRequired));
      });

      if (options.json) {
        console.log(JSON.stringify({
          pendingEvents: pendingEvents.map((e, i) => ({
            index: i + 1,
            id: e.Id,
            subject: e.Subject,
            start: e.Start.DateTime,
            end: e.End.DateTime,
            organizer: e.Organizer?.EmailAddress?.Name || e.Organizer?.EmailAddress?.Address,
            location: e.Location?.DisplayName,
          })),
        }, null, 2));
        return;
      }

      console.log('\nCalendar invitations awaiting your response:\n');
      console.log('\u2500'.repeat(60));

      if (pendingEvents.length === 0) {
        console.log('\n  No pending invitations found.\n');
        return;
      }

      for (let i = 0; i < pendingEvents.length; i++) {
        const event = pendingEvents[i];
        const dateStr = formatDate(event.Start.DateTime);
        const startTime = formatTime(event.Start.DateTime);
        const endTime = formatTime(event.End.DateTime);

        const response = getInvitationResponseStatus(event, userEmail) || 'None';
        const icon = getResponseIcon(response);

        console.log(`\n  [${i + 1}] ${icon} ${event.Subject}`);
        console.log(`      ${dateStr} ${startTime} - ${endTime}`);
        console.log(`      ID: ${event.Id}`);
        if (event.Location?.DisplayName) {
          console.log(`      Location: ${event.Location.DisplayName}`);
        }
        if (event.Organizer?.EmailAddress) {
          const org = event.Organizer.EmailAddress;
          console.log(`      Organizer: ${org.Name || org.Address}`);
        }
      }

      console.log('\n' + '\u2500'.repeat(60));
      console.log('\nTo respond, use:');
      console.log('  clippy respond accept --id <eventId>');
      console.log('  clippy respond decline --id <eventId>');
      console.log('  clippy respond tentative --id <eventId>');
      console.log('  clippy respond propose --id <eventId> --date 2026-07-08 --start 09:45 --end 10:15');
      console.log('');
      return;
    }

    if (!options.id) {
      writeError('Please specify the event id with --id.', options.json);
      if (!options.json) {
        console.error('Run `clippy respond list` to see pending invitations and IDs.');
      }
      process.exit(1);
    }

    const targetResult = await getCalendarEvent(authResult.token!, options.id);
    if (!targetResult.ok || !targetResult.data) {
      writeError(targetResult.error?.message || `Event not found: ${options.id}`, options.json);
      process.exit(1);
    }

    const targetEvent = targetResult.data;
    if (targetEvent.IsCancelled) {
      writeError(`Cannot respond to a cancelled event: ${targetEvent.Subject}`, options.json);
      process.exit(1);
    }
    if (targetEvent.IsOrganizer) {
      writeError('Cannot respond to an event you organized. Use update-event or delete-event instead.', options.json);
      process.exit(1);
    }

    let proposedWindow: { start: Date; end: Date } | undefined;
    if (actionLower === 'propose') {
      if (!options.start && !options.date) {
        writeError('Please specify a proposed time with --start, and optionally --date and --end.', options.json);
        process.exit(1);
      }
      if (!options.notify) {
        writeError('Proposing a new time requires notifying the organizer; omit --no-notify.', options.json);
        process.exit(1);
      }
      try {
        proposedWindow = resolveProposedResponseWindow(targetEvent, options);
      } catch (err) {
        writeError(err instanceof Error ? err.message : 'Invalid proposed time.', options.json);
        process.exit(1);
      }
    }

    if (!options.json) {
      console.log(`\nResponding to: ${targetEvent.Subject}`);
      console.log(`  ${formatDate(targetEvent.Start.DateTime)} ${formatTime(targetEvent.Start.DateTime)} - ${formatTime(targetEvent.End.DateTime)}`);
      console.log(`  Action: ${actionLower}`);
      if (proposedWindow) {
        console.log(`  Proposed: ${formatDate(proposedWindow.start.toISOString())} ${formatTime(proposedWindow.start.toISOString())} - ${formatTime(proposedWindow.end.toISOString())}`);
      }
      if (options.comment) {
        console.log(`  Comment: ${options.comment}`);
      }
      console.log('');
    }

    const response = await respondToEvent({
      token: authResult.token!,
      eventId: targetEvent.Id,
      response: actionLower as ResponseType,
      comment: options.comment,
      sendResponse: options.notify,
      proposedStart: proposedWindow?.start.toISOString(),
      proposedEnd: proposedWindow?.end.toISOString(),
    });

    if (!response.ok) {
      writeError(response.error?.message || 'Failed to respond', options.json);
      process.exit(1);
    }

    const actionPast = actionLower === 'tentative'
      ? 'tentatively accepted'
      : actionLower === 'propose'
        ? 'proposed a new time for'
        : `${actionLower}d`;
    if (options.json) {
      console.log(JSON.stringify({
        success: true,
        action: actionLower,
        proposedStart: proposedWindow?.start.toISOString(),
        proposedEnd: proposedWindow?.end.toISOString(),
        event: {
          id: targetEvent.Id,
          subject: targetEvent.Subject,
          start: targetEvent.Start.DateTime,
          end: targetEvent.End.DateTime,
        },
      }, null, 2));
    } else {
      console.log(`\u2713 Successfully ${actionPast} the invitation.`);
    }
  });
