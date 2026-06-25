import { afterEach, expect, test } from 'bun:test';
import {
  getCloudConfig,
  getEwsTokenScopes,
  getOAuthTokenEndpoint,
  resolveCloudEnvironment,
} from '../src/lib/cloud.js';
import { getCalendarEvent } from '../src/lib/ews-client.js';

const originalEnv = {
  CLIPPY_CLOUD: process.env.CLIPPY_CLOUD,
  EWS_AUTHORITY_HOST: process.env.EWS_AUTHORITY_HOST,
  EWS_ENDPOINT: process.env.EWS_ENDPOINT,
  EWS_RESOURCE: process.env.EWS_RESOURCE,
  EWS_TENANT_ID: process.env.EWS_TENANT_ID,
};
const originalFetch = globalThis.fetch;

afterEach(() => {
  for (const [key, value] of Object.entries(originalEnv)) {
    if (value === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = value;
    }
  }
  globalThis.fetch = originalFetch;
});

test('defaults to commercial Microsoft 365 endpoints', () => {
  delete process.env.CLIPPY_CLOUD;

  const config = getCloudConfig();

  expect(resolveCloudEnvironment()).toBe('commercial');
  expect(config.ewsEndpoint).toBe('https://outlook.office365.com/EWS/Exchange.asmx');
  expect(getOAuthTokenEndpoint(config)).toBe('https://login.microsoftonline.com/common/oauth2/v2.0/token');
  expect(getEwsTokenScopes(config)).toEqual([
    'https://outlook.office365.com/EWS.AccessAsUser.All offline_access',
    'https://outlook.office365.com/.default offline_access',
  ]);
});

test('selects US Government endpoints for gcc cloud', () => {
  process.env.CLIPPY_CLOUD = 'gcc';

  const config = getCloudConfig();

  expect(config.environment).toBe('gcc');
  expect(config.ewsEndpoint).toBe('https://outlook.office365.us/EWS/Exchange.asmx');
  expect(getOAuthTokenEndpoint(config)).toBe('https://login.microsoftonline.us/common/oauth2/v2.0/token');
  expect(getEwsTokenScopes(config)).toEqual([
    'https://outlook.office365.us/EWS.AccessAsUser.All offline_access',
    'https://outlook.office365.us/.default offline_access',
  ]);
});

test('allows explicit endpoint overrides', () => {
  process.env.CLIPPY_CLOUD = 'gcc';
  process.env.EWS_AUTHORITY_HOST = 'https://login.example.test/';
  process.env.EWS_RESOURCE = 'https://ews.example.test/';
  process.env.EWS_ENDPOINT = 'https://ews.example.test/EWS/Exchange.asmx';
  process.env.EWS_TENANT_ID = 'tenant-id';

  const config = getCloudConfig();

  expect(config.authorityHost).toBe('https://login.example.test');
  expect(config.ewsResource).toBe('https://ews.example.test');
  expect(config.ewsEndpoint).toBe('https://ews.example.test/EWS/Exchange.asmx');
  expect(getOAuthTokenEndpoint(config)).toBe('https://login.example.test/tenant-id/oauth2/v2.0/token');
  expect(getEwsTokenScopes(config)[0]).toBe('https://ews.example.test/EWS.AccessAsUser.All offline_access');
});

test('rejects unsupported cloud names', () => {
  process.env.CLIPPY_CLOUD = 'moon';

  expect(() => getCloudConfig()).toThrow('Unsupported CLIPPY_CLOUD "moon"');
});

test('EWS calls use the selected cloud endpoint', async () => {
  const urls: string[] = [];
  process.env.CLIPPY_CLOUD = 'gcc';

  globalThis.fetch = (async (input: RequestInfo | URL) => {
    urls.push(String(input));
    return new Response(`
      <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
        <soap:Body>
          <m:GetItemResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                             xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseMessages>
              <m:GetItemResponseMessage ResponseClass="Success">
                <m:ResponseCode>NoError</m:ResponseCode>
                <m:Items>
                  <t:CalendarItem>
                    <t:ItemId Id="event-id" />
                    <t:Subject>GCC event</t:Subject>
                    <t:Start>2026-06-08T15:30:00Z</t:Start>
                    <t:End>2026-06-08T16:15:00Z</t:End>
                  </t:CalendarItem>
                </m:Items>
              </m:GetItemResponseMessage>
            </m:ResponseMessages>
          </m:GetItemResponse>
        </soap:Body>
      </soap:Envelope>
    `, { status: 200 });
  }) as typeof fetch;

  const result = await getCalendarEvent('token', 'event-id');

  expect(result.ok).toBe(true);
  expect(urls).toEqual(['https://outlook.office365.us/EWS/Exchange.asmx']);
});
