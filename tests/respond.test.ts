import { expect, test } from 'bun:test';
import { getInvitationResponseStatus, isCancellationNoticeSubject, isPendingInvitation } from '../src/commands/respond.js';
import type { CalendarEvent } from '../src/lib/ews-client.js';

function calendarEvent(overrides: Partial<CalendarEvent>): CalendarEvent {
  return {
    Id: 'event-id',
    Subject: 'Test event',
    Start: { DateTime: '2026-06-08T15:30:00Z', TimeZone: 'UTC' },
    End: { DateTime: '2026-06-08T16:15:00Z', TimeZone: 'UTC' },
    ...overrides,
  };
}

test('pending invitation filter trusts MyResponseType when attendee records are absent', () => {
  const accepted = calendarEvent({
    MyResponseType: 'Accept',
    Attendees: undefined,
  });

  expect(getInvitationResponseStatus(accepted, 'andre.foeken@nedap.com')).toBe('Accepted');
  expect(isPendingInvitation(accepted, 'andre.foeken@nedap.com')).toBe(false);
});

test('pending invitation filter still includes unanswered MyResponseType values', () => {
  const unanswered = calendarEvent({
    MyResponseType: 'NoResponseReceived',
    Attendees: undefined,
  });

  expect(getInvitationResponseStatus(unanswered, 'andre.foeken@nedap.com')).toBe('NotResponded');
  expect(isPendingInvitation(unanswered, 'andre.foeken@nedap.com')).toBe(true);
});

test('pending invitation filter does not assume missing response metadata means pending', () => {
  const unknown = calendarEvent({
    MyResponseType: undefined,
    Attendees: undefined,
  });

  expect(getInvitationResponseStatus(unknown, 'andre.foeken@nedap.com')).toBeUndefined();
  expect(isPendingInvitation(unknown, 'andre.foeken@nedap.com')).toBe(false);
});

test('pending invitation filter excludes cancellation notices that Exchange does not mark cancelled', () => {
  const cancellationNotice = calendarEvent({
    Subject: 'Geannuleerd: TPV + connectors',
    MyResponseType: 'NoResponseReceived',
    IsCancelled: false,
  });

  expect(isCancellationNoticeSubject(cancellationNotice.Subject)).toBe(true);
  expect(isPendingInvitation(cancellationNotice, 'andre.foeken@nedap.com')).toBe(false);
});
