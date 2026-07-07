import { expect, test } from 'bun:test';
import {
  collectEventAvailabilityParticipants,
  evaluateSlotAvailability,
} from '../src/lib/calendar-availability.js';
import type { CalendarEvent, RawFreeBusyResult } from '../src/lib/ews-client.js';

function calendarEvent(overrides: Partial<CalendarEvent> = {}): CalendarEvent {
  return {
    Id: 'event-id',
    Subject: 'NTP Sync',
    Start: { DateTime: '2026-07-03T09:00:00Z', TimeZone: 'UTC' },
    End: { DateTime: '2026-07-03T10:00:00Z', TimeZone: 'UTC' },
    Organizer: {
      EmailAddress: {
        Name: 'Andre Foeken',
        Address: 'andre.foeken@nedap.com',
      },
    },
    Attendees: [
      {
        Type: 'Required',
        Status: { Response: 'None', Time: '' },
        EmailAddress: {
          Name: 'Anne Coppens',
          Address: 'anne.coppens@nedap.com',
        },
      },
      {
        Type: 'Optional',
        Status: { Response: 'None', Time: '' },
        EmailAddress: {
          Name: 'Andre Foeken',
          Address: 'andre.foeken@nedap.com',
        },
      },
    ],
    ...overrides,
  };
}

test('collectEventAvailabilityParticipants includes organizer and de-duplicates attendees', () => {
  expect(collectEventAvailabilityParticipants(calendarEvent())).toEqual([
    {
      email: 'andre.foeken@nedap.com',
      name: 'Andre Foeken',
      type: 'Organizer',
    },
    {
      email: 'anne.coppens@nedap.com',
      name: 'Anne Coppens',
      type: 'Required',
    },
  ]);
});

test('evaluateSlotAvailability ignores the current event while reporting other blockers', () => {
  const event = calendarEvent();
  const participants = collectEventAvailabilityParticipants(event);
  const rawFreeBusy: RawFreeBusyResult = {
    attendees: participants.map(participant => participant.email),
    intervalMinutes: 30,
    start: '2026-07-03T09:00:00Z',
    end: '2026-07-03T09:30:00Z',
    slots: [],
    attendeeDetails: [
      {
        email: 'andre.foeken@nedap.com',
        responseCode: 'NoError',
        events: [
          {
            start: '2026-07-03T09:00:00Z',
            end: '2026-07-03T10:00:00Z',
            busyType: 'Busy',
            subject: 'NTP Sync',
          },
        ],
      },
      {
        email: 'anne.coppens@nedap.com',
        responseCode: 'NoError',
        events: [
          {
            start: '2026-07-03T09:00:00Z',
            end: '2026-07-03T10:00:00Z',
            busyType: 'Busy',
            subject: 'NTP Sync',
          },
          {
            start: '2026-07-03T09:15:00Z',
            end: '2026-07-03T09:45:00Z',
            busyType: 'Tentative',
            subject: 'Another meeting',
          },
        ],
      },
    ],
  };

  const summary = evaluateSlotAvailability(
    rawFreeBusy,
    participants,
    '2026-07-03T09:00:00Z',
    '2026-07-03T09:30:00Z',
    event
  );

  expect(summary.allFree).toBe(false);
  expect(summary.participants[0]).toMatchObject({
    email: 'andre.foeken@nedap.com',
    available: true,
    status: 'Free',
    blockers: [],
  });
  expect(summary.participants[1]).toMatchObject({
    email: 'anne.coppens@nedap.com',
    available: false,
    status: 'Tentative',
  });
  expect(summary.participants[1].blockers).toHaveLength(1);
  expect(summary.participants[1].blockers[0].subject).toBe('Another meeting');
});

test('evaluateSlotAvailability ignores stale same-meeting copies and working-elsewhere entries', () => {
  const event = calendarEvent({
    End: { DateTime: '2026-07-03T09:30:00Z', TimeZone: 'UTC' },
  });
  const participants = collectEventAvailabilityParticipants(event);
  const rawFreeBusy: RawFreeBusyResult = {
    attendees: participants.map(participant => participant.email),
    intervalMinutes: 30,
    start: '2026-07-03T09:00:00Z',
    end: '2026-07-03T09:30:00Z',
    slots: [],
    attendeeDetails: participants.map(participant => ({
      email: participant.email,
      responseCode: 'NoError',
      events: [
        {
          start: '2026-07-03T00:00:00Z',
          end: '2026-07-04T00:00:00Z',
          busyType: 'WorkingElsewhere',
          subject: 'Remote',
        },
        {
          start: '2026-07-03T09:00:00Z',
          end: '2026-07-03T10:00:00Z',
          busyType: 'Tentative',
          subject: 'NTP Sync',
        },
      ],
    })),
  };

  const summary = evaluateSlotAvailability(
    rawFreeBusy,
    participants,
    '2026-07-03T09:00:00Z',
    '2026-07-03T09:30:00Z',
    event
  );

  expect(summary.allFree).toBe(true);
  expect(summary.participants.every(participant => participant.available)).toBe(true);
  expect(summary.participants.flatMap(participant => participant.blockers)).toEqual([]);
});
