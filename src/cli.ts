#!/usr/bin/env bun
import { Command } from 'commander';
import { whoamiCommand } from './commands/whoami.js';
import { calendarCommand } from './commands/calendar.js';
import { findtimeCommand } from './commands/findtime.js';
import { respondCommand } from './commands/respond.js';
import { createEventCommand } from './commands/create-event.js';
import { deleteEventCommand } from './commands/delete-event.js';
import { findCommand } from './commands/find.js';
import { freebusyCommand } from './commands/freebusy.js';
import { updateEventCommand } from './commands/update-event.js';
import { mailCommand } from './commands/mail.js';
import { foldersCommand } from './commands/folders.js';
import { sendCommand } from './commands/send.js';
import { draftsCommand } from './commands/drafts.js';

if (process.argv.includes('--read-only')) {
  process.env.CLIPPY_READONLY = 'true';
}
if (process.argv.includes('--gcc')) {
  process.env.CLIPPY_CLOUD = 'gcc';
}

const program = new Command();

program
  .name('clippy')
  .description('CLI for Microsoft 365/EWS')
  .version('0.1.0')
  .option('--gcc', 'Use Microsoft 365 US Government endpoints')
  .option('--read-only', 'Disable write commands');

program.addCommand(whoamiCommand);
program.addCommand(calendarCommand);
program.addCommand(findtimeCommand);
program.addCommand(respondCommand);
program.addCommand(createEventCommand);
program.addCommand(deleteEventCommand);
program.addCommand(findCommand);
program.addCommand(freebusyCommand);
program.addCommand(updateEventCommand);
program.addCommand(mailCommand);
program.addCommand(foldersCommand);
program.addCommand(sendCommand);
program.addCommand(draftsCommand);

program.parse();
