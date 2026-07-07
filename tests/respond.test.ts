import { expect, test } from 'bun:test';
import { getInvitationResponseStatus, isCancellationNoticeSubject, isPendingInvitation, resolveProposedResponseWindow } from '../src/commands/respond.js';
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

test('proposal window can move to a different date while preserving duration', () => {
  const window = resolveProposedResponseWindow(calendarEvent({}), {
    date: '2026-07-09',
    start: '09:45',
  });

  expect(window.start.getFullYear()).toBe(2026);
  expect(window.start.getMonth()).toBe(6);
  expect(window.start.getDate()).toBe(9);
  expect(window.start.getHours()).toBe(9);
  expect(window.start.getMinutes()).toBe(45);
  expect(window.end.getTime() - window.start.getTime()).toBe(45 * 60 * 1000);
});

test('proposal window can change duration with explicit end time', () => {
  const window = resolveProposedResponseWindow(calendarEvent({}), {
    date: '2026-07-09',
    start: '09:45',
    end: '10:15',
  });

  expect(window.start.getFullYear()).toBe(2026);
  expect(window.start.getMonth()).toBe(6);
  expect(window.start.getDate()).toBe(9);
  expect(window.start.getHours()).toBe(9);
  expect(window.start.getMinutes()).toBe(45);
  expect(window.end.getHours()).toBe(10);
  expect(window.end.getMinutes()).toBe(15);
  expect(window.end.getTime() - window.start.getTime()).toBe(30 * 60 * 1000);
});

test('proposal window can change duration with duration minutes', () => {
  const window = resolveProposedResponseWindow(calendarEvent({}), {
    date: '2026-07-09',
    start: '09:45',
    duration: '60',
  });

  expect(window.start.getHours()).toBe(9);
  expect(window.start.getMinutes()).toBe(45);
  expect(window.end.getTime() - window.start.getTime()).toBe(60 * 60 * 1000);
});

test('proposal window rejects conflicting end and duration options', () => {
  expect(() => resolveProposedResponseWindow(calendarEvent({}), {
    start: '09:45',
    end: '10:15',
    duration: '60',
  })).toThrow('Use only one of --end or --duration');
});
