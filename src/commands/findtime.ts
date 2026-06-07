import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getRawFreeBusy, getOwaUserInfo, type RawFreeBusySlot } from '../lib/ews-client.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

function formatDateTime(dateStr: string): string {
  const date = new Date(dateStr);
  return `${formatDate(dateStr)} ${formatTime(dateStr)}`;
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
    default:
      const parsed = new Date(day);
      return isNaN(parsed.getTime()) ? now : parsed;
  }
}

function getDateRange(startDay: string, endDay?: string): { start: Date; end: Date; label: string } {
  const now = new Date();

  switch (startDay.toLowerCase()) {
    case 'week':
    case 'thisweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      start.setDate(start.getDate() + (dayOfWeek === 0 ? 1 : 8 - dayOfWeek));
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

  const startDate = parseDay(startDay);
  startDate.setHours(0, 0, 0, 0);

  if (endDay) {
    const endDate = parseDay(endDay, startDate);
    endDate.setHours(23, 59, 59, 999);
    return {
      start: startDate,
      end: endDate,
      label: `${formatDate(startDate.toISOString())} - ${formatDate(endDate.toISOString())}`
    };
  }

  const endDate = new Date(startDate);
  endDate.setHours(23, 59, 59, 999);
  return { start: startDate, end: endDate, label: formatDate(startDate.toISOString()) };
}

function parseHour(value: string, name: string): number {
  const hour = Number(value);
  if (!Number.isInteger(hour) || hour < 0 || hour > 24) {
    throw new Error(`${name} must be an hour from 0 to 24.`);
  }
  return hour;
}

function parseDuration(value: string): number {
  const duration = Number(value);
  if (!Number.isInteger(duration) || duration <= 0) {
    throw new Error('--duration must be a positive whole number of minutes.');
  }
  if (duration % 5 !== 0) {
    throw new Error('--duration must be divisible by 5 minutes.');
  }
  return duration;
}

function availabilityInterval(durationMinutes: number): number {
  for (const interval of [30, 15, 10, 5]) {
    if (durationMinutes % interval === 0) return interval;
  }
  return 5;
}

function localMinutes(date: Date): number {
  return date.getHours() * 60 + date.getMinutes();
}

function isWithinWorkWindow(startStr: string, endStr: string, workStart: number, workEnd: number): boolean {
  const start = new Date(startStr);
  const end = new Date(endStr);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return false;
  if (start.toDateString() !== end.toDateString()) return false;
  return localMinutes(start) >= workStart * 60 && localMinutes(end) <= workEnd * 60;
}

function isContiguous(window: RawFreeBusySlot[]): boolean {
  for (let i = 1; i < window.length; i++) {
    if (new Date(window[i - 1].end).getTime() !== new Date(window[i].start).getTime()) {
      return false;
    }
  }
  return true;
}

function findAvailableWindows(
  slots: RawFreeBusySlot[],
  durationMinutes: number,
  intervalMinutes: number,
  workStart: number,
  workEnd: number
): Array<{ start: string; end: string }> {
  const requiredSlots = durationMinutes / intervalMinutes;
  const sortedSlots = [...slots].sort((a, b) => new Date(a.start).getTime() - new Date(b.start).getTime());
  const availableWindows: Array<{ start: string; end: string }> = [];

  for (let i = 0; i <= sortedSlots.length - requiredSlots; i++) {
    const window = sortedSlots.slice(i, i + requiredSlots);
    if (!window.every(slot => slot.allFree)) continue;
    if (!isContiguous(window)) continue;

    const start = window[0].start;
    const end = window[window.length - 1].end;
    if (!isWithinWorkWindow(start, end, workStart, workEnd)) continue;

    availableWindows.push({ start, end });
  }

  return availableWindows;
}

export const findtimeCommand = new Command('findtime')
  .description('Find available meeting times with one or more people')
  .argument('[start]', 'Start: today, tomorrow, monday-sunday, week, nextweek, or YYYY-MM-DD', 'nextweek')
  .argument('[endOrEmails...]', 'End day for range AND/OR email addresses')
  .option('--duration <minutes>', 'Meeting duration in minutes', '30')
  .option('--start <hour>', 'Work day start hour (0-23)', '9')
  .option('--end <hour>', 'Work day end hour (0-23)', '17')
  .option('--solo', 'Only check specified people, don\'t include yourself')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .action(async (startDay: string, endOrEmails: string[], options: {
    duration: string;
    start: string;
    end: string;
    solo?: boolean;
    json?: boolean;
    token?: string;
  }) => {
    let duration: number;
    let workStart: number;
    let workEnd: number;

    try {
      duration = parseDuration(options.duration);
      workStart = parseHour(options.start, '--start');
      workEnd = parseHour(options.end, '--end');
      if (workStart >= workEnd) {
        throw new Error('--start must be before --end.');
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Invalid findtime options.';
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

    // Parse arguments: figure out which are dates vs emails
    const dateKeywords = ['today', 'tomorrow', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday', 'week', 'thisweek', 'nextweek'];
    const isDateArg = (arg: string) => {
      if (dateKeywords.includes(arg.toLowerCase())) return true;
      if (/^\d{4}-\d{2}-\d{2}$/.test(arg)) return true;
      return false;
    };

    let endDay: string | undefined;
    let emails: string[] = [];

    for (const arg of endOrEmails) {
      if (arg.includes('@')) {
        emails.push(arg);
      } else if (isDateArg(arg) && !endDay) {
        endDay = arg;
      } else {
        emails.push(arg);
      }
    }

    // Get current user's email to include in search (unless --solo)
    if (!options.solo) {
      const userInfo = await getOwaUserInfo(authResult.token!);
      if (userInfo.ok && userInfo.data?.email) {
        // Add current user if not already in the list
        if (!emails.includes(userInfo.data.email)) {
          emails.unshift(userInfo.data.email);
        }
      }
    }

    if (emails.length === 0) {
      console.error('Error: Please provide at least one email address.');
      console.error('\nUsage: clippy findtime nextweek user@example.com');
      process.exit(1);
    }

    const { start, end, label } = getDateRange(startDay, endDay);
    const interval = availabilityInterval(duration);

    const result = await getRawFreeBusy(
      authResult.token!,
      emails,
      start.toISOString(),
      end.toISOString(),
      interval
    );

    if (!result.ok || !result.data) {
      if (options.json) {
        console.log(JSON.stringify({ error: result.error?.message || 'Failed to find meeting times' }, null, 2));
      } else {
        console.error(`Error: ${result.error?.message || 'Failed to find meeting times'}`);
      }
      process.exit(1);
    }

    const freeSlots = findAvailableWindows(result.data.slots, duration, interval, workStart, workEnd);

    if (options.json) {
      console.log(JSON.stringify({
        attendees: emails,
        duration: duration,
        dateRange: { start: start.toISOString(), end: end.toISOString() },
        availableSlots: freeSlots,
      }, null, 2));
      return;
    }

    console.log(`\n🗓️  Finding ${duration}-minute meeting times`);
    console.log(`   Attendees: ${emails.join(', ')}`);
    console.log(`   Date range: ${label}`);
    console.log('─'.repeat(50));

    if (freeSlots.length === 0) {
      console.log('\n  ❌ No available times found for all attendees.');
      console.log('     Try a longer date range or shorter meeting duration.');
    } else {
      console.log(`\n  ✅ Found ${freeSlots.length} available slot${freeSlots.length > 1 ? 's' : ''}:\n`);

      // Group by day
      const byDay = new Map<string, typeof freeSlots>();
      for (const slot of freeSlots) {
        const day = slot.start.split('T')[0];
        if (!byDay.has(day)) byDay.set(day, []);
        byDay.get(day)!.push(slot);
      }

      for (const [day, slots] of byDay) {
        const dayLabel = formatDate(new Date(day).toISOString());
        console.log(`  ${dayLabel}:`);
        for (const slot of slots) {
          console.log(`    🟢 ${formatTime(slot.start)} - ${formatTime(slot.end)}`);
        }
      }
    }
    console.log();
  });
