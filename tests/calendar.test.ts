import { expect, test } from 'bun:test';
import { formatAttendeePreview, limitEventAttendees } from '../src/commands/calendar.js';
import type { CalendarAttendee, CalendarEvent } from '../src/lib/ews-client.js';

function attendee(name: string, address = `${name.toLowerCase()}@nedap.com`): CalendarAttendee {
  return {
    Type: 'Required',
    Status: { Response: 'None', Time: '' },
    EmailAddress: { Name: name, Address: address },
  };
}

function event(attendees?: CalendarAttendee[]): CalendarEvent {
  return {
    Id: 'event-id',
    Subject: 'Test event',
    Start: { DateTime: '2026-07-20T09:00:00Z', TimeZone: 'UTC' },
    End: { DateTime: '2026-07-20T09:30:00Z', TimeZone: 'UTC' },
    Attendees: attendees,
  };
}

test('formatAttendeePreview keeps attendee output compact', () => {
  expect(formatAttendeePreview([attendee('Anne'), attendee('Bas'), attendee('Chris')], 2))
    .toBe('Anne, Bas +1 more');
  expect(formatAttendeePreview([attendee('Anne')], 1)).toBe('Anne');
});

test('limitEventAttendees preserves the original event while marking JSON previews', () => {
  const original = event([attendee('Anne'), attendee('Bas')]);
  const preview = limitEventAttendees(original, 1);

  expect(original.Attendees).toHaveLength(2);
  expect(preview.Attendees?.map(a => a.EmailAddress.Name)).toEqual(['Anne']);
  expect(preview.AttendeeCount).toBe(2);
  expect(preview.AttendeesTruncated).toBe(true);
});

test('limitEventAttendees leaves events without attendees unchanged', () => {
  const original = event();
  expect(limitEventAttendees(original, 1)).toBe(original);
});
