import { afterEach, expect, test } from 'bun:test';
import { createEvent, getCalendarEvent, updateEvent } from '../src/lib/ews-client.js';

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
});

function stubEws(envelopes: string[]) {
  globalThis.fetch = (async (_input: RequestInfo | URL, init?: RequestInit) => {
    envelopes.push(String(init?.body || ''));

    return new Response(`
      <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
        <soap:Body>
          <m:Response xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseMessages>
              <m:ResponseMessage ResponseClass="Success">
                <m:ResponseCode>NoError</m:ResponseCode>
                <m:Items>
                  <t:CalendarItem>
                    <t:ItemId Id="event-id" />
                    <t:Subject>Test event</t:Subject>
                    <t:Start>2026-06-08T15:30:00Z</t:Start>
                    <t:End>2026-06-08T16:15:00Z</t:End>
                    <t:MyResponseType>Accept</t:MyResponseType>
                    <t:Sensitivity>Private</t:Sensitivity>
                  </t:CalendarItem>
                </m:Items>
              </m:ResponseMessage>
            </m:ResponseMessages>
          </m:Response>
        </soap:Body>
      </soap:Envelope>
    `, { status: 200 });
  }) as typeof fetch;
}

test('createEvent writes reminder settings into the calendar item', async () => {
  const envelopes: string[] = [];
  stubEws(envelopes);

  const result = await createEvent({
    token: 'token',
    subject: 'Test event',
    start: '2026-06-08T15:30:00Z',
    end: '2026-06-08T16:15:00Z',
    reminderIsSet: true,
    reminderMinutesBeforeStart: 30,
  });

  expect(result.ok).toBe(true);
  expect(result.data?.reminderIsSet).toBe(true);
  expect(result.data?.reminderMinutesBeforeStart).toBe(30);
  expect(envelopes[0]).toContain('<t:ReminderIsSet>true</t:ReminderIsSet>');
  expect(envelopes[0]).toContain('<t:ReminderMinutesBeforeStart>30</t:ReminderMinutesBeforeStart>');
});

test('createEvent writes calendar sensitivity into the calendar item', async () => {
  const envelopes: string[] = [];
  stubEws(envelopes);

  const result = await createEvent({
    token: 'token',
    subject: 'Test event',
    start: '2026-06-08T15:30:00Z',
    end: '2026-06-08T16:15:00Z',
    sensitivity: 'Private',
  });

  expect(result.ok).toBe(true);
  expect(result.data?.Sensitivity).toBe('Private');
  expect(envelopes[0]).toContain('<t:Sensitivity>Private</t:Sensitivity>');
});

test('updateEvent writes reminder FieldURI updates', async () => {
  const envelopes: string[] = [];
  stubEws(envelopes);

  const result = await updateEvent({
    token: 'token',
    eventId: 'event-id',
    reminderIsSet: true,
    reminderMinutesBeforeStart: 30,
  });

  expect(result.ok).toBe(true);
  expect(result.data?.reminderIsSet).toBe(true);
  expect(result.data?.reminderMinutesBeforeStart).toBe(30);
  expect(envelopes[0]).toContain('FieldURI="item:ReminderIsSet"');
  expect(envelopes[0]).toContain('<t:ReminderIsSet>true</t:ReminderIsSet>');
  expect(envelopes[0]).toContain('FieldURI="item:ReminderMinutesBeforeStart"');
  expect(envelopes[0]).toContain('<t:ReminderMinutesBeforeStart>30</t:ReminderMinutesBeforeStart>');
  expect(envelopes[0]).toContain('SendMeetingInvitationsOrCancellations="SendToNone"');
});

test('updateEvent writes sensitivity FieldURI updates without meeting notifications', async () => {
  const envelopes: string[] = [];
  stubEws(envelopes);

  const result = await updateEvent({
    token: 'token',
    eventId: 'event-id',
    sensitivity: 'Private',
  });

  expect(result.ok).toBe(true);
  expect(result.data?.Sensitivity).toBe('Private');
  expect(envelopes[0]).toContain('FieldURI="item:Sensitivity"');
  expect(envelopes[0]).toContain('<t:Sensitivity>Private</t:Sensitivity>');
  expect(envelopes[0]).toContain('SendMeetingInvitationsOrCancellations="SendToNone"');
});

test('getCalendarEvent parses calendar sensitivity', async () => {
  const envelopes: string[] = [];
  stubEws(envelopes);

  const result = await getCalendarEvent('token', 'event-id');

  expect(result.ok).toBe(true);
  expect(result.data?.Sensitivity).toBe('Private');
  expect(result.data?.MyResponseType).toBe('Accept');
  expect(envelopes[0]).toContain('FieldURI="item:Sensitivity"');
  expect(envelopes[0]).toContain('FieldURI="calendar:MyResponseType"');
});

test('updateEvent can disable reminders without requiring reminder minutes', async () => {
  const envelopes: string[] = [];
  stubEws(envelopes);

  const result = await updateEvent({
    token: 'token',
    eventId: 'event-id',
    reminderIsSet: false,
  });

  expect(result.ok).toBe(true);
  expect(result.data?.reminderIsSet).toBe(false);
  expect(result.data?.reminderMinutesBeforeStart).toBeUndefined();
  expect(envelopes[0]).toContain('FieldURI="item:ReminderIsSet"');
  expect(envelopes[0]).toContain('<t:ReminderIsSet>false</t:ReminderIsSet>');
  expect(envelopes[0]).not.toContain('FieldURI="item:ReminderMinutesBeforeStart"');
});
