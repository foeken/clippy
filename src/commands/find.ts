import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { OUTLOOK_API } from '../lib/endpoints.js';

interface Person {
  DisplayName?: string;
  GivenName?: string;
  Surname?: string;
  JobTitle?: string;
  Department?: string;
  OfficeLocation?: string;
  UserPrincipalName?: string;
  ScoredEmailAddresses?: Array<{ Address?: string }>;
  Phones?: Array<{ Number?: string; Type?: string }>;
  PersonType?: { Class?: string; Subclass?: string };
}

async function searchPeople(
  token: string,
  query: string,
  filter?: 'people' | 'rooms'
): Promise<Person[]> {
  const url = `${OUTLOOK_API}/me/people?$search=${encodeURIComponent(query)}&$top=25`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
    },
  });

  if (!response.ok) {
    throw new Error(`Search failed: ${response.status}`);
  }

  const data = (await response.json()) as { value: Person[] };

  if (filter === 'rooms') {
    return data.value.filter(p => p.PersonType?.Subclass === 'Room');
  } else if (filter === 'people') {
    return data.value.filter(p => p.PersonType?.Subclass !== 'Room');
  }

  return data.value;
}

export const findCommand = new Command('find')
  .description('Search for people or rooms')
  .argument('<query>', 'Search query (name, email, etc.)')
  .option('--rooms', 'Only show rooms')
  .option('--people', 'Only show people (exclude rooms)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('-i, --interactive', 'Open browser to extract token automatically')
  .action(async (query: string, options: {
    rooms?: boolean;
    people?: boolean;
    json?: boolean;
    token?: string;
    interactive?: boolean;
  }) => {
    const authResult = await resolveAuth({
      token: options.token,
      interactive: options.interactive,
    });

    if (!authResult.success) {
      if (options.json) {
        console.log(JSON.stringify({ error: authResult.error }, null, 2));
      } else {
        console.error(`Error: ${authResult.error}`);
        console.error('\nRun `clippy login --interactive` to authenticate.');
      }
      process.exit(1);
    }

    const filter = options.rooms ? 'rooms' : options.people ? 'people' : undefined;

    try {
      const results = await searchPeople(authResult.token!, query, filter);

      if (options.json) {
        console.log(JSON.stringify({
          results: results.map(p => ({
            name: p.DisplayName,
            email: p.ScoredEmailAddresses?.[0]?.Address || p.UserPrincipalName,
            title: p.JobTitle,
            department: p.Department,
            type: p.PersonType?.Subclass === 'Room' ? 'Room' : 'Person',
          })),
        }, null, 2));
        return;
      }

      if (results.length === 0) {
        console.log(`\nNo results found for "${query}"\n`);
        return;
      }

      console.log(`\nSearch results for "${query}":\n`);
      console.log('\u2500'.repeat(60));

      for (const person of results) {
        const isRoom = person.PersonType?.Subclass === 'Room';
        const email = person.ScoredEmailAddresses?.[0]?.Address || person.UserPrincipalName;
        const icon = isRoom ? '\u{1F4CD}' : '\u{1F464}';

        console.log(`\n  ${icon} ${person.DisplayName}`);
        if (email) {
          console.log(`     ${email}`);
        }
        if (!isRoom) {
          if (person.JobTitle) {
            console.log(`     ${person.JobTitle}`);
          }
          if (person.Department) {
            console.log(`     ${person.Department}`);
          }
        }
      }

      console.log('\n' + '\u2500'.repeat(60) + '\n');
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Unknown error';
      if (options.json) {
        console.log(JSON.stringify({ error: message }, null, 2));
      } else {
        console.error(`Error: ${message}`);
      }
      process.exit(1);
    }
  });
