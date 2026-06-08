import { afterEach, expect, test } from 'bun:test';
import { createEvent, updateEvent } from '../src/lib/ews-client.js';

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
