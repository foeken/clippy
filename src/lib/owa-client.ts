import { OUTLOOK_BASE, OUTLOOK_API, GRAPH_BASE } from './endpoints.js';

export interface OwaRequestOptions {
  action: string;
  body: Record<string, unknown>;
  token: string;
}

export interface OwaError {
  code: string;
  message: string;
}

export interface OwaResponse<T = unknown> {
  ok: boolean;
  status: number;
  data?: T;
  error?: OwaError;
}

const USER_AGENT =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

export async function owaRequest<T = unknown>(
  options: OwaRequestOptions
): Promise<OwaResponse<T>> {
  const { action, body, token } = options;
  const url = `${OUTLOOK_BASE}/owa/service.svc?action=${action}&app=Mail&n=0`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as T;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface UserConfiguration {
  SessionSettings?: {
    UserDisplayName?: string;
    UserEmailAddress?: string;
  };
}

export async function getUserConfiguration(
  token: string
): Promise<OwaResponse<UserConfiguration>> {
  return owaRequest<UserConfiguration>({
    action: 'GetUserConfiguration',
    body: {
      __type: 'GetUserConfigurationRequest:#Exchange',
      Header: {
        __type: 'JsonRequestHeaders:#Exchange',
        RequestServerVersion: 'Exchange2016',
      },
      Body: {
        __type: 'GetUserConfigurationRequest:#Exchange',
        UserConfigurationName: {
          __type: 'UserConfigurationNameType:#Exchange',
          Name: 'OWA.SessionData',
        },
        UserConfigurationProperties: 'All',
      },
    },
    token,
  });
}

export interface OwaUserInfo {
  displayName: string;
  email: string;
}

export async function getOwaUserInfo(
  token: string
): Promise<OwaResponse<OwaUserInfo>> {
  // Use Outlook REST API to get user info
  const url = `${OUTLOOK_API}/me`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = await response.json() as { DisplayName?: string; EmailAddress?: string };

    return {
      ok: true,
      status: response.status,
      data: {
        displayName: data.DisplayName || 'Unknown',
        email: data.EmailAddress || 'Unknown',
      },
    };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface CalendarAttendee {
  Type: 'Required' | 'Optional' | 'Resource';
  Status: {
    Response: 'None' | 'Organizer' | 'TentativelyAccepted' | 'Accepted' | 'Declined' | 'NotResponded';
    Time: string;
  };
  EmailAddress: {
    Name: string;
    Address: string;
  };
}

export interface CalendarEvent {
  Id: string;
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  Location?: { DisplayName?: string };
  Organizer?: { EmailAddress?: { Name?: string; Address?: string } };
  Attendees?: CalendarAttendee[];
  IsAllDay?: boolean;
  IsCancelled?: boolean;
  IsOrganizer?: boolean;
  BodyPreview?: string;
  Categories?: string[];
  ShowAs?: string;
  Importance?: string;
  IsOnlineMeeting?: boolean;
  OnlineMeetingUrl?: string;
  WebLink?: string;
}

export interface CalendarViewResponse {
  value: CalendarEvent[];
}

export async function getCalendarEvents(
  token: string,
  startDateTime: string,
  endDateTime: string
): Promise<OwaResponse<CalendarEvent[]>> {
  const url = `${OUTLOOK_API}/me/calendarview?startDateTime=${encodeURIComponent(startDateTime)}&endDateTime=${encodeURIComponent(endDateTime)}&$orderby=Start/DateTime&$top=1000`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.timezone="Europe/Amsterdam"',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as CalendarViewResponse;

    return {
      ok: true,
      status: response.status,
      data: data.value,
    };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface FreeBusySlot {
  status: 'Free' | 'Busy' | 'Tentative';
  start: string;
  end: string;
  subject?: string;
}

export interface ScheduleInfo {
  scheduleId: string;
  availabilityView: string;
  scheduleItems: Array<{
    status: string;
    start: { dateTime: string; timeZone: string };
    end: { dateTime: string; timeZone: string };
    subject?: string;
    location?: string;
  }>;
}

/**
 * Get schedule/availability for multiple users using Microsoft Graph API.
 * Requires a Graph API token with Calendars.Read.Shared permission.
 */
export async function getScheduleForUsers(
  graphToken: string,
  emails: string[],
  startDateTime: string,
  endDateTime: string
): Promise<OwaResponse<ScheduleInfo[]>> {
  // Try getSchedule endpoint first
  const url = `${GRAPH_BASE}/v1.0/me/calendar/getSchedule`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${graphToken}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      body: JSON.stringify({
        schedules: emails,
        startTime: {
          dateTime: startDateTime,
          timeZone: 'Europe/Amsterdam',
        },
        endTime: {
          dateTime: endDateTime,
          timeZone: 'Europe/Amsterdam',
        },
        availabilityViewInterval: 30,
      }),
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({})) as { error?: { code?: string; message?: string } };
      return {
        ok: false,
        status: response.status,
        error: {
          code: errorData.error?.code || `HTTP_${response.status}`,
          message: errorData.error?.message || response.statusText,
        },
      };
    }

    const data = await response.json() as { value: ScheduleInfo[] };

    return {
      ok: true,
      status: response.status,
      data: data.value,
    };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Get schedule/availability for users using Outlook REST API.
 * Uses the same token as other Outlook API calls.
 */
export async function getScheduleViaOutlook(
  token: string,
  emails: string[],
  startDateTime: string,
  endDateTime: string,
  durationMinutes: number = 30
): Promise<OwaResponse<ScheduleInfo[]>> {
  // Try using FindMeetingTimes which can access other users' availability
  const url = `${OUTLOOK_API}/me/FindMeetingTimes`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
        Prefer: 'outlook.timezone="Europe/Amsterdam"',
      },
      body: JSON.stringify({
        Attendees: emails.map(email => ({
          EmailAddress: { Address: email, Name: email },
          Type: 'Required',
        })),
        TimeConstraint: {
          Timeslots: [{
            Start: { DateTime: startDateTime, TimeZone: 'W. Europe Standard Time' },
            End: { DateTime: endDateTime, TimeZone: 'W. Europe Standard Time' },
          }],
        },
        MeetingDuration: `PT${durationMinutes}M`,
        ReturnSuggestionReasons: true,
        MinimumAttendeePercentage: 100,
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = await response.json() as {
      MeetingTimeSuggestions?: Array<{
        MeetingTimeSlot: {
          Start: { DateTime: string; TimeZone: string };
          End: { DateTime: string; TimeZone: string };
        };
        Confidence: number;
        AttendeeAvailability?: Array<{
          Attendee: { EmailAddress: { Address: string } };
          Availability: string;
        }>;
      }>;
    };

    // Transform FindMeetingTimes response to our format
    const schedules: ScheduleInfo[] = emails.map(email => ({
      scheduleId: email,
      availabilityView: '',
      scheduleItems: [],
    }));

    // Parse meeting suggestions to find free/busy times
    if (data.MeetingTimeSuggestions && data.MeetingTimeSuggestions.length > 0) {
      // Build free slots from suggestions
      const freeSlots = data.MeetingTimeSuggestions.map(s => ({
        start: s.MeetingTimeSlot.Start.DateTime,
        end: s.MeetingTimeSlot.End.DateTime,
      }));

      for (const schedule of schedules) {
        // Add free slots
        schedule.scheduleItems = freeSlots.map(slot => ({
          status: 'Free',
          start: { dateTime: slot.start, timeZone: 'W. Europe Standard Time' },
          end: { dateTime: slot.end, timeZone: 'W. Europe Standard Time' },
        }));
      }
    } else {
      // No meeting times found - users are busy for the entire period
      for (const schedule of schedules) {
        schedule.scheduleItems = [{
          status: 'Busy',
          start: { dateTime: startDateTime, timeZone: 'W. Europe Standard Time' },
          end: { dateTime: endDateTime, timeZone: 'W. Europe Standard Time' },
          subject: 'No available times',
        }];
      }
    }

    return {
      ok: true,
      status: response.status,
      data: schedules,
    };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Get free/busy info for current user by analyzing their calendar events.
 * Note: Looking up other users requires Microsoft Graph API with different permissions.
 */
export async function getFreeBusy(
  token: string,
  startDateTime: string,
  endDateTime: string
): Promise<OwaResponse<FreeBusySlot[]>> {
  // Use calendar view to get events, then convert to free/busy
  const result = await getCalendarEvents(token, startDateTime, endDateTime);

  if (!result.ok || !result.data) {
    return {
      ok: false,
      status: result.status,
      error: result.error,
    };
  }

  const slots: FreeBusySlot[] = result.data
    .filter(event => !event.IsCancelled)
    .map(event => ({
      status: event.ShowAs === 'Free' ? 'Free' as const :
              event.ShowAs === 'Tentative' ? 'Tentative' as const : 'Busy' as const,
      start: event.Start.DateTime,
      end: event.End.DateTime,
      subject: event.Subject,
    }));

  return {
    ok: true,
    status: 200,
    data: slots,
  };
}

export interface RecurrencePattern {
  Type: 'Daily' | 'Weekly' | 'AbsoluteMonthly' | 'RelativeMonthly' | 'AbsoluteYearly' | 'RelativeYearly';
  Interval: number;  // How many days/weeks/months between occurrences
  DaysOfWeek?: string[];  // For weekly: ['Monday', 'Wednesday', 'Friday']
  DayOfMonth?: number;  // For monthly: 15 (15th of month)
  Month?: number;  // For yearly: 1-12
  Index?: 'First' | 'Second' | 'Third' | 'Fourth' | 'Last';  // For relative patterns
}

export interface RecurrenceRange {
  Type: 'EndDate' | 'NoEnd' | 'Numbered';
  StartDate: string;  // YYYY-MM-DD
  EndDate?: string;  // For EndDate type
  NumberOfOccurrences?: number;  // For Numbered type
}

export interface Recurrence {
  Pattern: RecurrencePattern;
  Range: RecurrenceRange;
}

export interface CreateEventOptions {
  token: string;
  subject: string;
  start: string;  // ISO datetime
  end: string;    // ISO datetime
  body?: string;
  location?: string;
  attendees?: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' | 'Resource' }>;
  isOnlineMeeting?: boolean;
  recurrence?: Recurrence;
}

export interface CreatedEvent {
  Id: string;
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  WebLink?: string;
  OnlineMeetingUrl?: string;
}

/**
 * Create a new calendar event.
 */
export async function createEvent(
  options: CreateEventOptions
): Promise<OwaResponse<CreatedEvent>> {
  const { token, subject, start, end, body, location, attendees, isOnlineMeeting, recurrence } = options;
  const url = `${OUTLOOK_API}/me/events`;

  const eventBody: Record<string, unknown> = {
    Subject: subject,
    Start: {
      DateTime: start,
      TimeZone: 'Europe/Amsterdam',
    },
    End: {
      DateTime: end,
      TimeZone: 'Europe/Amsterdam',
    },
  };

  if (body) {
    eventBody.Body = {
      ContentType: 'Text',
      Content: body,
    };
  }

  if (location) {
    eventBody.Location = {
      DisplayName: location,
    };
  }

  if (attendees && attendees.length > 0) {
    eventBody.Attendees = attendees.map(a => ({
      EmailAddress: {
        Address: a.email,
        Name: a.name || a.email,
      },
      Type: a.type || 'Required',
    }));
  }

  if (isOnlineMeeting) {
    eventBody.IsOnlineMeeting = true;
    eventBody.OnlineMeetingProvider = 'TeamsForBusiness';
  }

  if (recurrence) {
    eventBody.Recurrence = recurrence;
  }

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.timezone="Europe/Amsterdam"',
      },
      body: JSON.stringify(eventBody),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = (await response.json()) as CreatedEvent;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface UpdateEventOptions {
  token: string;
  eventId: string;
  subject?: string;
  start?: string;  // ISO datetime
  end?: string;    // ISO datetime
  body?: string;
  location?: string;
  attendees?: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' | 'Resource' }>;
  isOnlineMeeting?: boolean;
}

/**
 * Update an existing calendar event.
 */
export async function updateEvent(
  options: UpdateEventOptions
): Promise<OwaResponse<CreatedEvent>> {
  const { token, eventId, subject, start, end, body, location, attendees, isOnlineMeeting } = options;
  const url = `${OUTLOOK_API}/me/events/${encodeURIComponent(eventId)}`;

  const eventBody: Record<string, unknown> = {};

  if (subject !== undefined) {
    eventBody.Subject = subject;
  }

  if (start !== undefined) {
    eventBody.Start = {
      DateTime: start,
      TimeZone: 'Europe/Amsterdam',
    };
  }

  if (end !== undefined) {
    eventBody.End = {
      DateTime: end,
      TimeZone: 'Europe/Amsterdam',
    };
  }

  if (body !== undefined) {
    eventBody.Body = {
      ContentType: 'Text',
      Content: body,
    };
  }

  if (location !== undefined) {
    eventBody.Location = {
      DisplayName: location,
    };
  }

  if (attendees !== undefined) {
    eventBody.Attendees = attendees.map(a => ({
      EmailAddress: {
        Address: a.email,
        Name: a.name || a.email,
      },
      Type: a.type || 'Required',
    }));
  }

  if (isOnlineMeeting !== undefined) {
    eventBody.IsOnlineMeeting = isOnlineMeeting;
    if (isOnlineMeeting) {
      eventBody.OnlineMeetingProvider = 'TeamsForBusiness';
    }
  }

  try {
    const response = await fetch(url, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.timezone="Europe/Amsterdam"',
      },
      body: JSON.stringify(eventBody),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = (await response.json()) as CreatedEvent;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface Room {
  Address: string;
  Name: string;
}

export interface RoomList {
  Address: string;
  Name: string;
}

/**
 * Get available room lists (buildings/locations).
 */
export async function getRoomLists(
  token: string
): Promise<OwaResponse<RoomList[]>> {
  // Try Graph API first (works with Outlook token in some cases)
  const urls = [
    `${GRAPH_BASE}/v1.0/places/microsoft.graph.roomList`,
    `${OUTLOOK_API}/me/findRoomLists`,
  ];

  for (const url of urls) {
    try {
      const response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'User-Agent': USER_AGENT,
          Accept: 'application/json',
        },
      });

      if (response.ok) {
        const data = (await response.json()) as { value: Array<{ emailAddress?: string; address?: string; displayName?: string; Name?: string; Address?: string }> };
        const rooms = data.value.map(r => ({
          Address: r.emailAddress || r.address || r.Address || '',
          Name: r.displayName || r.Name || '',
        }));
        if (rooms.length > 0) {
          return { ok: true, status: response.status, data: rooms };
        }
      }
    } catch {
      // Try next URL
    }
  }

  return {
    ok: false,
    status: 404,
    error: {
      code: 'NOT_FOUND',
      message: 'No room lists found',
    },
  };
}

/**
 * Get rooms in a room list or all rooms.
 */
export async function getRooms(
  token: string,
  roomListAddress?: string
): Promise<OwaResponse<Room[]>> {
  // Try multiple endpoints
  const urls = roomListAddress
    ? [
        `${GRAPH_BASE}/v1.0/places/${encodeURIComponent(roomListAddress)}/microsoft.graph.roomList/rooms`,
        `${OUTLOOK_API}/me/findRooms(RoomList='${encodeURIComponent(roomListAddress)}')`,
      ]
    : [
        `${GRAPH_BASE}/v1.0/places/microsoft.graph.room`,
        `${OUTLOOK_API}/me/findRooms`,
      ];

  for (const url of urls) {
    try {
      const response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'User-Agent': USER_AGENT,
          Accept: 'application/json',
        },
      });

      if (response.ok) {
        const data = (await response.json()) as { value: Array<{ emailAddress?: string; address?: string; displayName?: string; Name?: string; Address?: string }> };
        const rooms = data.value.map(r => ({
          Address: r.emailAddress || r.address || r.Address || '',
          Name: r.displayName || r.Name || '',
        }));
        if (rooms.length > 0) {
          return { ok: true, status: response.status, data: rooms };
        }
      }
    } catch {
      // Try next URL
    }
  }

  return {
    ok: false,
    status: 404,
    error: {
      code: 'NOT_FOUND',
      message: 'No rooms found',
    },
  };
}

/**
 * Search for rooms/resources using the People API.
 */
export async function searchRooms(
  token: string,
  query: string = 'room'
): Promise<OwaResponse<Room[]>> {
  // Use People search API with room filter
  const searchQuery = query || 'room';
  const url = `${OUTLOOK_API}/me/people?$search=${encodeURIComponent(searchQuery)}&$top=50`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as {
      value: Array<{
        DisplayName?: string;
        ScoredEmailAddresses?: Array<{ Address?: string }>;
        PersonType?: { Class?: string; Subclass?: string };
      }>;
    };

    // Filter to only rooms (PersonType.Subclass === 'Room')
    const rooms: Room[] = data.value
      .filter(p => p.PersonType?.Subclass === 'Room')
      .map(p => ({
        Name: p.DisplayName || '',
        Address: p.ScoredEmailAddresses?.[0]?.Address || '',
      }))
      .filter(r => r.Address);

    return { ok: true, status: response.status, data: rooms };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Delete a calendar event (without sending cancellation notices).
 */
export async function deleteEvent(
  token: string,
  eventId: string
): Promise<OwaResponse<void>> {
  const url = `${OUTLOOK_API}/me/events/${encodeURIComponent(eventId)}`;

  try {
    const response = await fetch(url, {
      method: 'DELETE',
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    return { ok: true, status: response.status };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Cancel a calendar event (sends cancellation notices to attendees).
 * Use this when you're the organizer and want to notify attendees.
 */
export async function cancelEvent(
  token: string,
  eventId: string,
  comment?: string
): Promise<OwaResponse<void>> {
  const url = `${OUTLOOK_API}/me/events/${encodeURIComponent(eventId)}/cancel`;

  try {
    const body: Record<string, unknown> = {};
    if (comment) {
      body.Comment = comment;
    }

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    return { ok: true, status: response.status };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

// Email types
export interface EmailAddress {
  Name?: string;
  Address?: string;
}

export interface EmailMessage {
  Id: string;
  Subject?: string;
  BodyPreview?: string;
  Body?: {
    ContentType: string;
    Content: string;
  };
  From?: {
    EmailAddress?: EmailAddress;
  };
  ToRecipients?: Array<{ EmailAddress?: EmailAddress }>;
  CcRecipients?: Array<{ EmailAddress?: EmailAddress }>;
  ReceivedDateTime?: string;
  SentDateTime?: string;
  IsRead?: boolean;
  IsDraft?: boolean;
  HasAttachments?: boolean;
  Importance?: 'Low' | 'Normal' | 'High';
  Flag?: {
    FlagStatus?: 'NotFlagged' | 'Flagged' | 'Complete';
  };
}

export interface EmailListResponse {
  value: EmailMessage[];
  '@odata.nextLink'?: string;
}

export interface GetEmailsOptions {
  token: string;
  folder?: string;  // inbox, sentitems, drafts, deleteditems, archive, junkemail
  top?: number;
  skip?: number;
  filter?: string;
  search?: string;
  select?: string[];
  orderBy?: string;
}

/**
 * Get emails from a folder.
 */
export async function getEmails(
  options: GetEmailsOptions
): Promise<OwaResponse<EmailListResponse>> {
  const {
    token,
    folder = 'inbox',
    top = 10,
    skip = 0,
    filter,
    search,
    select = ['Id', 'Subject', 'BodyPreview', 'From', 'ReceivedDateTime', 'IsRead', 'HasAttachments', 'Importance', 'Flag'],
    orderBy = 'ReceivedDateTime desc',
  } = options;

  const params = new URLSearchParams();
  params.set('$top', top.toString());
  if (skip > 0) params.set('$skip', skip.toString());
  if (filter) params.set('$filter', filter);
  if (search) params.set('$search', `"${search}"`);
  params.set('$select', select.join(','));
  // Note: $orderby is ignored when $search is used (results are ranked by relevance)
  if (!search) params.set('$orderby', orderBy);

  const url = `${OUTLOOK_API}/me/mailfolders/${folder}/messages?${params}`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.body-content-type="text"',
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = (await response.json()) as EmailListResponse;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Get a single email by ID.
 */
export async function getEmail(
  token: string,
  messageId: string
): Promise<OwaResponse<EmailMessage>> {
  const url = `${OUTLOOK_API}/me/messages/${encodeURIComponent(messageId)}`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.body-content-type="text"',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as EmailMessage;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

// Attachment types
export interface Attachment {
  Id: string;
  Name: string;
  ContentType: string;
  Size: number;
  IsInline: boolean;
  ContentId?: string;
  ContentBytes?: string;  // Base64 encoded content
}

export interface AttachmentListResponse {
  value: Attachment[];
}

/**
 * Get list of attachments for an email.
 */
export async function getAttachments(
  token: string,
  messageId: string
): Promise<OwaResponse<AttachmentListResponse>> {
  const url = `${OUTLOOK_API}/me/messages/${encodeURIComponent(messageId)}/attachments`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as AttachmentListResponse;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Get a single attachment with content.
 */
export async function getAttachment(
  token: string,
  messageId: string,
  attachmentId: string
): Promise<OwaResponse<Attachment>> {
  const url = `${OUTLOOK_API}/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as Attachment;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Update email properties (read status, flag, etc.)
 */
export async function updateEmail(
  token: string,
  messageId: string,
  updates: {
    IsRead?: boolean;
    Flag?: {
      FlagStatus: 'NotFlagged' | 'Flagged' | 'Complete';
    };
  }
): Promise<OwaResponse<EmailMessage>> {
  const url = `${OUTLOOK_API}/me/messages/${encodeURIComponent(messageId)}`;

  try {
    const response = await fetch(url, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify(updates),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = (await response.json()) as EmailMessage;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Move an email to a different folder.
 */
export async function moveEmail(
  token: string,
  messageId: string,
  destinationFolder: string
): Promise<OwaResponse<EmailMessage>> {
  const url = `${OUTLOOK_API}/me/messages/${encodeURIComponent(messageId)}/move`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify({
        DestinationId: destinationFolder,
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = (await response.json()) as EmailMessage;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

// Folder types
export interface MailFolder {
  Id: string;
  DisplayName: string;
  ParentFolderId?: string;
  ChildFolderCount: number;
  UnreadItemCount: number;
  TotalItemCount: number;
}

export interface MailFolderListResponse {
  value: MailFolder[];
}

/**
 * Get list of mail folders.
 */
export async function getMailFolders(
  token: string,
  parentFolderId?: string
): Promise<OwaResponse<MailFolderListResponse>> {
  const url = parentFolderId
    ? `${OUTLOOK_API}/me/mailfolders/${encodeURIComponent(parentFolderId)}/childfolders`
    : `${OUTLOOK_API}/me/mailfolders?$top=100`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as MailFolderListResponse;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Create a new mail folder.
 */
export async function createMailFolder(
  token: string,
  displayName: string,
  parentFolderId?: string
): Promise<OwaResponse<MailFolder>> {
  const url = parentFolderId
    ? `${OUTLOOK_API}/me/mailfolders/${encodeURIComponent(parentFolderId)}/childfolders`
    : `${OUTLOOK_API}/me/mailfolders`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify({ DisplayName: displayName }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = (await response.json()) as MailFolder;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Update (rename) a mail folder.
 */
export async function updateMailFolder(
  token: string,
  folderId: string,
  displayName: string
): Promise<OwaResponse<MailFolder>> {
  const url = `${OUTLOOK_API}/me/mailfolders/${encodeURIComponent(folderId)}`;

  try {
    const response = await fetch(url, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify({ DisplayName: displayName }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = (await response.json()) as MailFolder;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Delete a mail folder.
 */
export async function deleteMailFolder(
  token: string,
  folderId: string
): Promise<OwaResponse<void>> {
  const url = `${OUTLOOK_API}/me/mailfolders/${encodeURIComponent(folderId)}`;

  try {
    const response = await fetch(url, {
      method: 'DELETE',
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    return { ok: true, status: response.status };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface EmailAttachment {
  name: string;
  contentType: string;
  contentBytes: string;  // Base64 encoded
}

/**
 * Send a new email, optionally with attachments.
 */
export async function sendEmail(
  token: string,
  options: {
    to: string[];
    cc?: string[];
    bcc?: string[];
    subject: string;
    body: string;
    bodyType?: 'Text' | 'HTML';
    attachments?: EmailAttachment[];
  }
): Promise<OwaResponse<void>> {
  const message: Record<string, unknown> = {
    Subject: options.subject,
    Body: {
      ContentType: options.bodyType || 'Text',
      Content: options.body,
    },
    ToRecipients: options.to.map(email => ({
      EmailAddress: { Address: email },
    })),
  };

  if (options.cc && options.cc.length > 0) {
    message.CcRecipients = options.cc.map(email => ({
      EmailAddress: { Address: email },
    }));
  }

  if (options.bcc && options.bcc.length > 0) {
    message.BccRecipients = options.bcc.map(email => ({
      EmailAddress: { Address: email },
    }));
  }

  // If no attachments, use simple sendmail
  if (!options.attachments || options.attachments.length === 0) {
    const url = `${OUTLOOK_API}/me/sendmail`;

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
          'User-Agent': USER_AGENT,
          Accept: 'application/json',
        },
        body: JSON.stringify({ Message: message }),
      });

      if (!response.ok) {
        const errorText = await response.text();
        return {
          ok: false,
          status: response.status,
          error: {
            code: `HTTP_${response.status}`,
            message: errorText || response.statusText,
          },
        };
      }

      return { ok: true, status: response.status };
    } catch (err) {
      return {
        ok: false,
        status: 0,
        error: {
          code: 'NETWORK_ERROR',
          message: err instanceof Error ? err.message : 'Unknown error',
        },
      };
    }
  }

  // With attachments, use draft workflow: create draft -> add attachments -> send
  try {
    // Step 1: Create draft
    const createUrl = `${OUTLOOK_API}/me/messages`;
    const createResponse = await fetch(createUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify(message),
    });

    if (!createResponse.ok) {
      const errorText = await createResponse.text();
      return {
        ok: false,
        status: createResponse.status,
        error: {
          code: `HTTP_${createResponse.status}`,
          message: errorText || createResponse.statusText,
        },
      };
    }

    const draft = await createResponse.json() as { Id: string };
    const draftId = draft.Id;

    // Step 2: Add attachments
    for (const attachment of options.attachments) {
      const attachUrl = `${OUTLOOK_API}/me/messages/${encodeURIComponent(draftId)}/attachments`;
      const attachResponse = await fetch(attachUrl, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
          'User-Agent': USER_AGENT,
          Accept: 'application/json',
        },
        body: JSON.stringify({
          '@odata.type': '#Microsoft.OutlookServices.FileAttachment',
          Name: attachment.name,
          ContentType: attachment.contentType,
          ContentBytes: attachment.contentBytes,
        }),
      });

      if (!attachResponse.ok) {
        const errorText = await attachResponse.text();
        return {
          ok: false,
          status: attachResponse.status,
          error: {
            code: `HTTP_${attachResponse.status}`,
            message: `Failed to add attachment ${attachment.name}: ${errorText || attachResponse.statusText}`,
          },
        };
      }
    }

    // Step 3: Send the draft
    const sendUrl = `${OUTLOOK_API}/me/messages/${encodeURIComponent(draftId)}/send`;
    const sendResponse = await fetch(sendUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
      },
    });

    if (!sendResponse.ok) {
      const errorText = await sendResponse.text();
      return {
        ok: false,
        status: sendResponse.status,
        error: {
          code: `HTTP_${sendResponse.status}`,
          message: errorText || sendResponse.statusText,
        },
      };
    }

    return { ok: true, status: sendResponse.status };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Create or send a reply to an email.
 * Always uses draft approach for proper control over formatting and quote separation.
 */
async function createReplyDraft(
  token: string,
  messageId: string,
  comment: string,
  replyAll: boolean = false,
  isHtml: boolean = false
): Promise<OwaResponse<{ draftId: string }>> {
  const createAction = replyAll ? 'createreplyall' : 'createreply';
  const createUrl = `${OUTLOOK_API}/me/messages/${encodeURIComponent(messageId)}/${createAction}`;

  try {
    // Step 1: Create reply draft (gets us the quoted original)
    const createResponse = await fetch(createUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.body-content-type="html"',
      },
      body: JSON.stringify({}),
    });

    if (!createResponse.ok) {
      const errorText = await createResponse.text();
      return {
        ok: false,
        status: createResponse.status,
        error: {
          code: `HTTP_${createResponse.status}`,
          message: errorText || createResponse.statusText,
        },
      };
    }

    const draft = await createResponse.json() as { Id: string; Body?: { Content?: string } };
    const draftId = draft.Id;
    const quotedOriginal = draft.Body?.Content || '';

    // Sanitize any separator/MIME boundary artifacts (Outlook sometimes injects them in text replies)
    // Also remove <hr> tags which Outlook uses as separators (converts to ____ in text mode)
    const sanitizedOriginal = quotedOriginal
      .replace(/<hr[^>]*>/gi, '')  // Remove <hr> separator tags
      .replace(/<[^>]*>\s*_{5,}\s*<\/[^>]*>/gi, '')
      .replace(/(?:&#95;|&lowbar;){5,}/gi, '')
      .replace(/_{5,}/g, '')
      .replace(/^\s*--[-A-Za-z0-9_]+\s*$/gm, '')
      .replace(/^\s*Content-(Type|Transfer-Encoding):.*$/gmi, '')
      .replace(/^\s*charset=.*$/gmi, '')
      .replace(/^\s*Content-Id:.*$/gmi, '')
      .replace(/^\s*Content-Description:.*$/gmi, '')
      .replace(/^\s*Content-Disposition:.*$/gmi, '')
      .replace(/\n{3,}/g, '\n\n');

    // Step 2: Update draft with our HTML prepended to the quoted content
    const updateUrl = `${OUTLOOK_API}/me/messages/${encodeURIComponent(draftId)}`;

    // Prepare the reply content
    let replyContent: string;
    if (isHtml) {
      // Extract just the body content if comment is a full HTML document
      const bodyMatch = comment.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
      replyContent = bodyMatch ? bodyMatch[1].trim() : comment;
    } else {
      // Convert plain text to HTML - escape HTML entities and convert newlines to <br>
      replyContent = comment
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\n/g, '<br>');
    }

    // Combine our HTML with the quoted original
    // Outlook puts the quote header in a div with id="divRplyFwdMsg", we insert before it
    let combinedBody: string;
    const replyDiv = `<div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; line-height: 1.5; margin-bottom: 20px;">${replyContent}</div>`;

    if (sanitizedOriginal.includes('divRplyFwdMsg')) {
      // Insert our content before the reply/forward message div
      combinedBody = sanitizedOriginal.replace(
        /(<div[^>]*id=["']?divRplyFwdMsg["']?)/i,
        `${replyDiv}$1`
      );
    } else if (sanitizedOriginal.includes('<body')) {
      // Fallback: insert after <body> tag
      combinedBody = sanitizedOriginal.replace(
        /(<body[^>]*>)/i,
        `$1${replyDiv}`
      );
    } else {
      // No body tag, just prepend
      combinedBody = `${replyDiv}${sanitizedOriginal}`;
    }

    const updateResponse = await fetch(updateUrl, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify({
        Body: {
          ContentType: 'HTML',
          Content: combinedBody,
        },
      }),
    });

    if (!updateResponse.ok) {
      const errorText = await updateResponse.text();
      return {
        ok: false,
        status: updateResponse.status,
        error: {
          code: `HTTP_${updateResponse.status}`,
          message: errorText || updateResponse.statusText,
        },
      };
    }

    return { ok: true, status: updateResponse.status, data: { draftId } };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export async function replyToEmailDraft(
  token: string,
  messageId: string,
  comment: string,
  replyAll: boolean = false,
  isHtml: boolean = false
): Promise<OwaResponse<{ draftId: string }>> {
  return createReplyDraft(token, messageId, comment, replyAll, isHtml);
}

export async function replyToEmail(
  token: string,
  messageId: string,
  comment: string,
  replyAll: boolean = false,
  isHtml: boolean = false
): Promise<OwaResponse<void>> {
  const draftResult = await createReplyDraft(token, messageId, comment, replyAll, isHtml);

  if (!draftResult.ok || !draftResult.data) {
    return draftResult as OwaResponse<void>;
  }

  const draftId = draftResult.data.draftId;

  // Send the draft
  const sendUrl = `${OUTLOOK_API}/me/messages/${encodeURIComponent(draftId)}/send`;
  try {
    const sendResponse = await fetch(sendUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
      },
    });

    if (!sendResponse.ok) {
      const errorText = await sendResponse.text();
      return {
        ok: false,
        status: sendResponse.status,
        error: {
          code: `HTTP_${sendResponse.status}`,
          message: errorText || sendResponse.statusText,
        },
      };
    }

    return { ok: true, status: sendResponse.status };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Forward an email to one or more recipients.
 */
export async function forwardEmail(
  token: string,
  messageId: string,
  toRecipients: string[],
  comment?: string
): Promise<OwaResponse<void>> {
  const url = `${OUTLOOK_API}/me/messages/${encodeURIComponent(messageId)}/forward`;

  try {
    const body: Record<string, unknown> = {
      ToRecipients: toRecipients.map(email => ({
        EmailAddress: { Address: email },
      })),
    };

    if (comment) {
      body.Comment = comment;
    }

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    return { ok: true, status: response.status };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export type ResponseType = 'accept' | 'decline' | 'tentative';

export interface RespondToEventOptions {
  token: string;
  eventId: string;
  response: ResponseType;
  comment?: string;
  sendResponse?: boolean;
}

/**
 * Respond to a calendar event (accept, decline, or tentatively accept).
 */
export async function respondToEvent(
  options: RespondToEventOptions
): Promise<OwaResponse<void>> {
  const { token, eventId, response, comment, sendResponse = true } = options;

  const actionMap: Record<ResponseType, string> = {
    accept: 'accept',
    decline: 'decline',
    tentative: 'tentativelyAccept',
  };

  const action = actionMap[response];
  const url = `${OUTLOOK_API}/me/events/${encodeURIComponent(eventId)}/${action}`;

  try {
    const body: Record<string, unknown> = {
      SendResponse: sendResponse,
    };

    if (comment) {
      body.Comment = comment;
    }

    const res = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const errorText = await res.text();
      return {
        ok: false,
        status: res.status,
        error: {
          code: `HTTP_${res.status}`,
          message: errorText || res.statusText,
        },
      };
    }

    return { ok: true, status: res.status };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Get a single calendar event by ID.
 */
export async function getCalendarEvent(
  token: string,
  eventId: string
): Promise<OwaResponse<CalendarEvent>> {
  const url = `${OUTLOOK_API}/me/events/${encodeURIComponent(eventId)}`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.timezone="Europe/Amsterdam"',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as CalendarEvent;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export async function validateSession(token: string): Promise<boolean> {
  // Use Outlook REST API to validate the token
  const url = `${OUTLOOK_API}/me/mailfolders/inbox`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
      },
    });

    // 200 means valid session, 401/403 means expired/invalid
    return response.ok;
  } catch {
    return false;
  }
}
