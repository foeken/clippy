import { afterEach, expect, test } from 'bun:test';
import { updateEvent } from '../src/lib/ews-client.js';

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
                    <t:Subject>Jeroen - Andre</t:Subject>
                    <t:Start>2026-06-10T07:00:00Z</t:Start>
                    <t:End>2026-06-10T07:20:00Z</t:End>
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

test('updateEvent writes a local title subject update without meeting notifications', async () => {
  const envelopes: string[] = [];
  stubEws(envelopes);

  const result = await updateEvent({
    token: 'token',
    eventId: 'event-id',
    subject: 'Jeroen - Andre',
  });

  expect(result.ok).toBe(true);
  expect(result.data?.Subject).toBe('Jeroen - Andre');
  expect(envelopes[0]).toContain('FieldURI="item:Subject"');
  expect(envelopes[0]).toContain('<t:Subject>Jeroen - Andre</t:Subject>');
  expect(envelopes[0]).toContain('SendMeetingInvitationsOrCancellations="SendToNone"');
});
