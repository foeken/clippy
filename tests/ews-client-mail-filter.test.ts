import { afterEach, expect, test } from 'bun:test';
import { getEmails } from '../src/lib/ews-client.js';

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
});

function stubFindItem(envelopes: string[]) {
  globalThis.fetch = (async (_input: RequestInfo | URL, init?: RequestInit) => {
    envelopes.push(String(init?.body || ''));

    return new Response(`
      <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
        <soap:Body>
          <m:FindItemResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                              xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseMessages>
              <m:FindItemResponseMessage ResponseClass="Success">
                <m:ResponseCode>NoError</m:ResponseCode>
                <m:RootFolder>
                  <t:Items>
                    <t:Message>
                      <t:ItemId Id="flagged-message-id" />
                      <t:Subject>Flagged message</t:Subject>
                      <t:Flag>
                        <t:FlagStatus>Flagged</t:FlagStatus>
                      </t:Flag>
                    </t:Message>
                  </t:Items>
                </m:RootFolder>
              </m:FindItemResponseMessage>
            </m:ResponseMessages>
          </m:FindItemResponse>
        </soap:Body>
      </soap:Envelope>
    `, { status: 200 });
  }) as typeof fetch;
}

test('getEmails filters flagged mail with the scalar MAPI flag status property', async () => {
  const envelopes: string[] = [];
  stubFindItem(envelopes);

  const result = await getEmails({
    token: 'token',
    filter: "Flag/FlagStatus eq 'Flagged'",
  });

  expect(result.ok).toBe(true);
  expect(result.data?.value[0].Flag?.FlagStatus).toBe('Flagged');
  expect(envelopes[0]).toContain('<t:ExtendedFieldURI PropertyTag="0x1090" PropertyType="Integer" />');
  expect(envelopes[0]).toContain('<t:Constant Value="2" />');
  expect(envelopes[0]).not.toContain('item:Flag/FlagStatus');
});
