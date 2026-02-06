#!/usr/bin/env bun
export {};

// Pre-parse --gcc flag before modules load (endpoints.ts reads env at import time)
if (process.argv.includes('--gcc')) {
  process.env.CLIPPY_CLOUD = 'gcc';
}

const { Command } = await import('commander');
const { loginCommand } = await import('./commands/login.js');
const { whoamiCommand } = await import('./commands/whoami.js');
const { calendarCommand } = await import('./commands/calendar.js');
const { findtimeCommand } = await import('./commands/findtime.js');
const { respondCommand } = await import('./commands/respond.js');
const { createEventCommand } = await import('./commands/create-event.js');
const { deleteEventCommand } = await import('./commands/delete-event.js');
const { findCommand } = await import('./commands/find.js');
const { updateEventCommand } = await import('./commands/update-event.js');
const { refreshCommand } = await import('./commands/refresh.js');
const { keepaliveCommand } = await import('./commands/keepalive.js');
const { mailCommand } = await import('./commands/mail.js');
const { foldersCommand } = await import('./commands/folders.js');
const { sendCommand } = await import('./commands/send.js');
const { draftsCommand } = await import('./commands/drafts.js');
const { CLOUD_ENV } = await import('./lib/endpoints.js');

const program = new Command();

program
  .name('clippy')
  .description('CLI for Microsoft 365/OWA')
  .version('0.1.0')
  .option('--gcc', 'Use Office 365 US Government (GCC) endpoints');

program.addCommand(loginCommand);
program.addCommand(whoamiCommand);
program.addCommand(calendarCommand);
program.addCommand(findtimeCommand);
program.addCommand(respondCommand);
program.addCommand(createEventCommand);
program.addCommand(deleteEventCommand);
program.addCommand(findCommand);
program.addCommand(updateEventCommand);
program.addCommand(refreshCommand);
program.addCommand(keepaliveCommand);
program.addCommand(mailCommand);
program.addCommand(foldersCommand);
program.addCommand(sendCommand);
program.addCommand(draftsCommand);

if (CLOUD_ENV === 'gcc') {
  console.log('[cloud: GCC (outlook.office365.us)]');
}

program.parse();
