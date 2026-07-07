import type {
  CalendarEvent,
  RawFreeBusyEvent,
  RawFreeBusyResult,
} from './ews-client.js';

export type AvailabilityParticipantType = 'Organizer' | 'Required' | 'Optional' | 'Resource';

export interface AvailabilityParticipant {
  email: string;
  name?: string;
  type: AvailabilityParticipantType;
}

export interface SlotAvailabilityParticipant extends AvailabilityParticipant {
  available: boolean;
  status: 'Free' | 'Tentative' | 'Busy' | 'OOF' | 'NoData';
  blockers: RawFreeBusyEvent[];
  responseCode?: string;
}

export interface SlotAvailabilitySummary {
  start: string;
  end: string;
  allFree: boolean;
  participants: SlotAvailabilityParticipant[];
}

export function collectEventAvailabilityParticipants(
  event: CalendarEvent,
  fallbackOrganizerEmail?: string
): AvailabilityParticipant[] {
  const participants: AvailabilityParticipant[] = [];
  const seen = new Set<string>();

  const add = (participant: AvailabilityParticipant) => {
    const email = participant.email.trim();
    if (!email) return;

    const key = email.toLowerCase();
    if (seen.has(key)) return;

    seen.add(key);
    participants.push({ ...participant, email });
  };

  const organizer = event.Organizer?.EmailAddress;
  add({
    email: organizer?.Address || fallbackOrganizerEmail || '',
    name: organizer?.Name,
    type: 'Organizer',
  });

  for (const attendee of event.Attendees || []) {
    const type = attendee.Type === 'Optional'
      ? 'Optional'
      : attendee.Type === 'Resource'
        ? 'Resource'
        : 'Required';

    add({
      email: attendee.EmailAddress?.Address || '',
      name: attendee.EmailAddress?.Name,
      type,
    });
  }

  return participants;
}

export function evaluateSlotAvailability(
  result: RawFreeBusyResult,
  participants: AvailabilityParticipant[],
  proposedStart: string,
  proposedEnd: string,
  currentEvent?: CalendarEvent
): SlotAvailabilitySummary {
  const start = new Date(proposedStart);
  const end = new Date(proposedEnd);

  const byEmail = new Map(
    (result.attendeeDetails || []).map(details => [details.email.toLowerCase(), details])
  );

  const fallbackStatuses = new Map<string, { status: SlotAvailabilityParticipant['status']; blockers: RawFreeBusyEvent[] }>();
  for (const slot of result.slots) {
    if (!dateRangesOverlap(start, end, new Date(slot.start), new Date(slot.end))) continue;

    for (const attendee of slot.attendees) {
      const key = attendee.email.toLowerCase();
      const existing = fallbackStatuses.get(key);
      fallbackStatuses.set(key, {
        status: strongestStatus(existing?.status, attendee.status),
        blockers: [
          ...(existing?.blockers || []),
          ...(attendee.events || []),
        ],
      });
    }
  }

  const participantSummaries = participants.map(participant => {
    const details = byEmail.get(participant.email.toLowerCase());
    const blockers = details
      ? details.events.filter(event =>
        isBlockingBusyType(event.busyType) &&
        eventOverlapsWindow(event, start, end) &&
        !isCurrentEventBlocker(event, currentEvent)
      )
      : (fallbackStatuses.get(participant.email.toLowerCase())?.blockers || [])
        .filter(event => isBlockingBusyType(event.busyType) && !isCurrentEventBlocker(event, currentEvent));

    const hasNoData = details
      ? details.responseCode !== undefined && details.responseCode !== 'NoError'
      : fallbackStatuses.get(participant.email.toLowerCase())?.status === 'NoData';

    const status = hasNoData
      ? 'NoData'
      : blockers.length > 0
        ? blockers.reduce<SlotAvailabilityParticipant['status']>(
          (status, blocker) => strongestStatus(status, normalizeBusyType(blocker.busyType)),
          'Free'
        )
        : 'Free';

    return {
      ...participant,
      available: status === 'Free',
      status,
      blockers,
      ...(details?.responseCode ? { responseCode: details.responseCode } : {}),
    };
  });

  return {
    start: start.toISOString(),
    end: end.toISOString(),
    allFree: participantSummaries.every(participant => participant.available),
    participants: participantSummaries,
  };
}

export function availabilityIntervalMinutes(start: Date, end: Date): number {
  const durationMinutes = Math.max(5, Math.round((end.getTime() - start.getTime()) / 60000));
  for (const interval of [60, 30, 15, 10, 5]) {
    if (durationMinutes % interval === 0) return interval;
  }
  return 5;
}

function normalizeBusyType(value?: string): SlotAvailabilityParticipant['status'] {
  switch (value) {
    case 'Tentative':
      return 'Tentative';
    case 'Busy':
      return 'Busy';
    case 'OOF':
      return 'OOF';
    case 'WorkingElsewhere':
      return 'Free';
    case 'Free':
      return 'Free';
    default:
      return 'NoData';
  }
}

function isBlockingBusyType(value?: string): boolean {
  return value !== 'Free' && value !== 'WorkingElsewhere';
}

function strongestStatus(
  current: SlotAvailabilityParticipant['status'] | undefined,
  next: SlotAvailabilityParticipant['status']
): SlotAvailabilityParticipant['status'] {
  const priority: Record<SlotAvailabilityParticipant['status'], number> = {
    Free: 0,
    Tentative: 1,
    Busy: 2,
    OOF: 3,
    NoData: 4,
  };

  if (!current) return next;
  return priority[next] > priority[current] ? next : current;
}

function eventOverlapsWindow(event: RawFreeBusyEvent, start: Date, end: Date): boolean {
  return dateRangesOverlap(start, end, new Date(event.start), new Date(event.end));
}

function dateRangesOverlap(startA: Date, endA: Date, startB: Date, endB: Date): boolean {
  if (
    Number.isNaN(startA.getTime()) ||
    Number.isNaN(endA.getTime()) ||
    Number.isNaN(startB.getTime()) ||
    Number.isNaN(endB.getTime())
  ) {
    return false;
  }

  return startA < endB && endA > startB;
}

function isCurrentEventBlocker(event: RawFreeBusyEvent, currentEvent?: CalendarEvent): boolean {
  if (!currentEvent) return false;

  const currentStart = new Date(currentEvent.Start.DateTime);
  const currentEnd = new Date(currentEvent.End.DateTime);
  const eventStart = new Date(event.start);
  const eventEnd = new Date(event.end);

  if (
    Number.isNaN(currentStart.getTime()) ||
    Number.isNaN(currentEnd.getTime()) ||
    Number.isNaN(eventStart.getTime()) ||
    Number.isNaN(eventEnd.getTime())
  ) {
    return false;
  }

  const currentSubject = currentEvent.Subject.trim().toLowerCase();
  const blockerSubject = event.subject?.trim().toLowerCase();

  if (blockerSubject && blockerSubject !== currentSubject) return false;

  return dateRangesOverlap(eventStart, eventEnd, currentStart, currentEnd);
}
