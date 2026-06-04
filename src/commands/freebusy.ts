import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getOwaUserInfo, getRawFreeBusy, type RawFreeBusySlot } from '../lib/ews-client.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

function parseDay(day: string, baseDate: Date = new Date()): Date {
  const now = new Date(baseDate);

  switch (day.toLowerCase()) {
    case 'today':
      return now;
    case 'tomorrow':
      now.setDate(now.getDate() + 1);
      return now;
    case 'monday':
    case 'tuesday':
    case 'wednesday':
    case 'thursday':
    case 'friday':
    case 'saturday':
    case 'sunday': {
      const days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
      const targetDay = days.indexOf(day.toLowerCase());
      const currentDay = now.getDay();
      let diff = targetDay - currentDay;
      if (diff <= 0) diff += 7;
      now.setDate(now.getDate() + diff);
      return now;
    }
    default: {
      const parsed = new Date(day);
      return isNaN(parsed.getTime()) ? now : parsed;
    }
  }
}

function getDateRange(startDay: string, endDay?: string): { start: Date; end: Date; label: string } {
  const now = new Date();

  switch (startDay.toLowerCase()) {
    case 'week':
    case 'thisweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      start.setDate(start.getDate() + diff);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 4);
      end.setHours(23, 59, 59, 999);
      return { start, end, label: 'This Week (Mon-Fri)' };
    }
    case 'nextweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const daysUntilNextMonday = dayOfWeek === 0 ? 1 : 8 - dayOfWeek;
      start.setDate(start.getDate() + daysUntilNextMonday);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 4);
      end.setHours(23, 59, 59, 999);
      return { start, end, label: 'Next Week (Mon-Fri)' };
    }
  }

  const start = parseDay(startDay);
  start.setHours(0, 0, 0, 0);

  if (endDay) {
    const end = parseDay(endDay, start);
    end.setHours(23, 59, 59, 999);
    return { start, end, label: `${formatDate(start.toISOString())} - ${formatDate(end.toISOString())}` };
  }

  const end = new Date(start);
  end.setHours(23, 59, 59, 999);
  return { start, end, label: formatDate(start.toISOString()) };
}

function isDateArg(arg: string): boolean {
  const dateKeywords = ['today', 'tomorrow', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday', 'week', 'thisweek', 'nextweek'];
  return dateKeywords.includes(arg.toLowerCase()) || /^\d{4}-\d{2}-\d{2}$/.test(arg);
}

function parseHour(value: string, name: string): number {
  const hour = Number.parseInt(value, 10);
  if (!Number.isInteger(hour) || hour < 0 || hour > 24) {
    throw new Error(`${name} must be an hour from 0 to 24.`);
  }
  return hour;
}

function parseInterval(value: string): number {
  const interval = Number.parseInt(value, 10);
  const allowed = [5, 10, 15, 30, 60];
  if (!allowed.includes(interval)) {
    throw new Error(`--interval must be one of: ${allowed.join(', ')}.`);
  }
  return interval;
}

function isWeekend(dateStr: string): boolean {
  const day = new Date(dateStr).getDay();
  return day === 0 || day === 6;
}

function isWithinWorkHours(dateStr: string, workStart: number, workEnd: number): boolean {
  const hour = new Date(dateStr).getHours();
  return hour >= workStart && hour < workEnd;
}

function filterSlots(slots: RawFreeBusySlot[], options: { workStart: number; workEnd: number; includeWeekends?: boolean }): RawFreeBusySlot[] {
  return slots.filter(slot => {
    if (!options.includeWeekends && isWeekend(slot.start)) return false;
    return isWithinWorkHours(slot.start, options.workStart, options.workEnd);
  });
}

function statusLine(slot: RawFreeBusySlot): string {
  return slot.attendees
    .map(attendee => `${attendee.email}: ${attendee.status}`)
    .join(', ');
}

function printSlots(slots: RawFreeBusySlot[], showAll: boolean): void {
  const byDay = new Map<string, RawFreeBusySlot[]>();

  for (const slot of slots) {
    const day = new Date(slot.start).toDateString();
    if (!byDay.has(day)) byDay.set(day, []);
    byDay.get(day)!.push(slot);
  }

  for (const [day, daySlots] of byDay) {
    console.log(`\n${formatDate(new Date(day).toISOString())}:`);
    for (const slot of daySlots) {
      const prefix = slot.allFree ? 'FREE' : 'BUSY';
      const details = showAll ? `  ${statusLine(slot)}` : '';
      console.log(`  ${prefix} ${formatTime(slot.start)} - ${formatTime(slot.end)}${details}`);
    }
  }
}

export const freebusyCommand = new Command('freebusy')
  .description('Show raw Exchange free/busy slots for one or more people')
  .argument('[start]', 'Start: today, tomorrow, monday-sunday, week, nextweek, or YYYY-MM-DD', 'nextweek')
  .argument('[endOrEmails...]', 'End day for range AND/OR email addresses')
  .option('--interval <minutes>', 'Slot size in minutes: 5, 10, 15, 30, or 60', '30')
  .option('--work-start <hour>', 'Work day start hour (0-24)', '9')
  .option('--work-end <hour>', 'Work day end hour (0-24)', '17')
  .option('--include-weekends', 'Include Saturday and Sunday slots')
  .option('--all', 'Show all raw slots, including busy/tentative/OOF slots')
  .option('--solo', 'Only check specified people, do not include yourself')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (startDay: string, endOrEmails: string[], options: {
    interval: string;
    workStart: string;
    workEnd: string;
    includeWeekends?: boolean;
    all?: boolean;
    solo?: boolean;
    json?: boolean;
    token?: string;
  }) => {
    let interval: number;
    let workStart: number;
    let workEnd: number;

    try {
      interval = parseInterval(options.interval);
      workStart = parseHour(options.workStart, '--work-start');
      workEnd = parseHour(options.workEnd, '--work-end');
      if (workStart >= workEnd) {
        throw new Error('--work-start must be before --work-end.');
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Invalid free/busy options.';
      if (options.json) {
        console.log(JSON.stringify({ error: message }, null, 2));
      } else {
        console.error(`Error: ${message}`);
      }
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

    let endDay: string | undefined;
    const emails: string[] = [];

    for (const arg of endOrEmails) {
      if (arg.includes('@')) {
        emails.push(arg);
      } else if (isDateArg(arg) && !endDay) {
        endDay = arg;
      } else {
        emails.push(arg);
      }
    }

    if (!options.solo) {
      const userInfo = await getOwaUserInfo(authResult.token!);
      if (userInfo.ok && userInfo.data?.email && !emails.includes(userInfo.data.email)) {
        emails.unshift(userInfo.data.email);
      }
    }

    if (emails.length === 0) {
      const message = 'Please provide at least one email address, or omit --solo to include yourself.';
      if (options.json) {
        console.log(JSON.stringify({ error: message }, null, 2));
      } else {
        console.error(`Error: ${message}`);
        console.error('\nUsage: clippy freebusy nextweek user@example.com');
      }
      process.exit(1);
    }

    const { start, end, label } = getDateRange(startDay, endDay);
    const result = await getRawFreeBusy(
      authResult.token!,
      emails,
      start.toISOString(),
      end.toISOString(),
      interval
    );

    if (!result.ok || !result.data) {
      const message = result.error?.message || 'Failed to fetch raw free/busy data';
      if (options.json) {
        console.log(JSON.stringify({ error: message }, null, 2));
      } else {
        console.error(`Error: ${message}`);
      }
      process.exit(1);
    }

    const filteredSlots = filterSlots(result.data.slots, {
      workStart,
      workEnd,
      includeWeekends: options.includeWeekends,
    });
    const availableSlots = filteredSlots.filter(slot => slot.allFree);
    const outputSlots = options.all ? filteredSlots : availableSlots;

    if (options.json) {
      console.log(JSON.stringify({
        attendees: result.data.attendees,
        intervalMinutes: result.data.intervalMinutes,
        dateRange: { start: result.data.start, end: result.data.end },
        filters: {
          workStart,
          workEnd,
          includeWeekends: Boolean(options.includeWeekends),
          all: Boolean(options.all),
        },
        availableSlots,
        slots: outputSlots,
      }, null, 2));
      return;
    }

    console.log(`\nRaw free/busy (${interval}-minute slots)`);
    console.log(`Attendees: ${result.data.attendees.join(', ')}`);
    console.log(`Date range: ${label}`);
    console.log(`Window: ${workStart}:00 - ${workEnd}:00${options.includeWeekends ? '' : ' weekdays only'}`);

    if (outputSlots.length === 0) {
      console.log(options.all
        ? '\nNo slots found in the selected window.'
        : '\nNo all-free slots found in the selected window. Use --all to inspect busy/tentative/OOF statuses.');
      console.log('');
      return;
    }

    printSlots(outputSlots, Boolean(options.all));
    console.log('');
  });
