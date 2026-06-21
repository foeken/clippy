import { afterEach, expect, test } from 'bun:test';
import { getEmails, moveEmail, moveEmailToFolderId } from '../src/lib/ews-client.js';

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
});

function stubEws(envelopes: string[]) {
  globalThis.fetch = (async (_input: RequestInfo | URL, init?: RequestInit) => {
    const envelope = String(init?.body || '');
    envelopes.push(envelope);

    if (envelope.includes('<m:FindFolder')) {
      return new Response(`
        <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
          <soap:Body>
            <m:FindFolderResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                                  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
              <m:ResponseMessages>
                <m:FindFolderResponseMessage ResponseClass="Success">
                  <m:ResponseCode>NoError</m:ResponseCode>
                  <m:RootFolder>
                    <t:Folders>
                      <t:Folder>
                        <t:FolderId Id="archive-folder-id" />
                        <t:DisplayName>Archive</t:DisplayName>
                        <t:ChildFolderCount>0</t:ChildFolderCount>
                        <t:UnreadCount>0</t:UnreadCount>
                        <t:TotalCount>0</t:TotalCount>
                      </t:Folder>
                    </t:Folders>
                  </m:RootFolder>
                </m:FindFolderResponseMessage>
              </m:ResponseMessages>
            </m:FindFolderResponse>
          </soap:Body>
        </soap:Envelope>
      `, { status: 200 });
    }

    if (envelope.includes('<m:FindItem')) {
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
                        <t:ItemId Id="archived-message-id" />
                        <t:Subject>Archived message</t:Subject>
                      </t:Message>
                    </t:Items>
                  </m:RootFolder>
                </m:FindItemResponseMessage>
              </m:ResponseMessages>
            </m:FindItemResponse>
          </soap:Body>
        </soap:Envelope>
      `, { status: 200 });
    }

    return new Response(`
      <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
        <soap:Body>
          <m:MoveItemResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                              xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseMessages>
              <m:MoveItemResponseMessage ResponseClass="Success">
                <m:ResponseCode>NoError</m:ResponseCode>
                <m:Items>
                  <t:Message>
                    <t:ItemId Id="moved-message-id" />
                  </t:Message>
                </m:Items>
              </m:MoveItemResponseMessage>
            </m:ResponseMessages>
          </m:MoveItemResponse>
        </soap:Body>
      </soap:Envelope>
    `, { status: 200 });
  }) as typeof fetch;
}

test('getEmails resolves archive by concrete folder id', async () => {
  const envelopes: string[] = [];
  stubEws(envelopes);

  const result = await getEmails({ token: 'token', folder: 'archive' });

  expect(result.ok).toBe(true);
  expect(result.data?.value[0].Id).toBe('archived-message-id');
  expect(envelopes[0]).toContain('<m:FindFolder');
  expect(envelopes[1]).toContain('<m:FindItem');
  expect(envelopes[1]).toContain('<t:FolderId Id="archive-folder-id" />');
  expect(envelopes[1]).not.toContain('<t:DistinguishedFolderId');
});

test('moveEmail resolves archive folder name by concrete folder id', async () => {
  const envelopes: string[] = [];
  stubEws(envelopes);

  const result = await moveEmail('token', 'message-id', 'archive');

  expect(result.ok).toBe(true);
  expect(result.data?.Id).toBe('moved-message-id');
  expect(envelopes[0]).toContain('<m:FindFolder');
  expect(envelopes[1]).toContain('<m:MoveItem>');
  expect(envelopes[1]).toContain('<t:FolderId Id="archive-folder-id" />');
  expect(envelopes[1]).not.toContain('<t:DistinguishedFolderId');
});

test('moveEmailToFolderId emits a concrete FolderId destination', async () => {
  const envelopes: string[] = [];
  stubEws(envelopes);

  const result = await moveEmailToFolderId('token', 'message-id', 'archive-folder-id');

  expect(result.ok).toBe(true);
  expect(result.data?.Id).toBe('moved-message-id');
  expect(envelopes[0]).toContain('<t:FolderId Id="archive-folder-id" />');
  expect(envelopes[0]).not.toContain('<t:DistinguishedFolderId');
});
