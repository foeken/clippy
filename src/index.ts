// Library exports for programmatic usage
export { resolveAuth } from './lib/auth.js';
export type { AuthResult } from './lib/auth.js';

export {
  validateSession, getOwaUserInfo,
  getCalendarEvents, getCalendarEvent, createEvent, updateEvent, deleteEvent, cancelEvent, respondToEvent,
  getEmails, getEmail, sendEmail, replyToEmail, replyToEmailDraft, forwardEmail, updateEmail, moveEmail, moveEmailToFolderId,
  createDraft, updateDraft, sendDraftById, deleteDraftById, addAttachmentToDraft,
  getMailFolders, createMailFolder, updateMailFolder, deleteMailFolder,
  getAttachments, getAttachment,
  resolveNames, getRoomLists, getRooms, searchRooms,
  getScheduleViaOutlook, getRawFreeBusy, getFreeBusy,
} from './lib/ews-client.js';

export type {
  OwaResponse, OwaError, OwaUserInfo,
  CalendarEvent, CalendarAttendee, CreatedEvent, CreateEventOptions, UpdateEventOptions,
  CalendarShowAs, CalendarSensitivity,
  EmailMessage, EmailListResponse, GetEmailsOptions, EmailAttachment,
  MailFolder, MailFolderListResponse,
  Attachment, AttachmentListResponse,
  Room, RoomList, ScheduleInfo, FreeBusySlot,
  MergedFreeBusyStatus, RawFreeBusyAttendeeStatus, RawFreeBusySlot, RawFreeBusyResult,
  Recurrence, RecurrencePattern, RecurrenceRange,
  ResponseType, RespondToEventOptions,
} from './lib/ews-client.js';
