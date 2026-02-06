# Clippy

A command-line interface for Microsoft 365 / Outlook Web Access (OWA). Manage your calendar and email directly from the terminal.

## Installation

```bash
# Clone the repository
git clone https://github.com/foeken/clippy.git
cd clippy

# Install dependencies
bun install

# Run directly
bun run src/cli.ts <command>

# Or link globally
bun link
clippy <command>
```

## Configuration

Copy `.env.example` to `.env` and adjust for your environment:

```bash
cp .env.example .env
```

| Variable | Values | Default | Description |
|---|---|---|---|
| `CLIPPY_CLOUD` | `commercial`, `gcc` | `commercial` | Microsoft cloud environment |
| `CLIPPY_TOKEN` | Bearer token string | — | Skip browser login with a token directly |

### Cloud Environments

- **`commercial`** — Standard Microsoft 365 (`outlook.office.com`)
- **`gcc`** — Office 365 US Government (`outlook.office365.us`)

You can also pass `--gcc` on any command instead of using the `.env` file:

```bash
clippy --gcc login --interactive
clippy --gcc calendar
```

---

## Authentication

Clippy uses browser-based authentication to obtain tokens from Outlook. During login, it captures both an access token (short-lived, ~1 hour) and a refresh token (long-lived, days/weeks).

```bash
# First time setup - opens browser for login
clippy login --interactive

# Check who you're logged in as
clippy whoami

# Refresh token (also runs automatically in background)
clippy refresh
```

### How Token Refresh Works

1. **Initial login**: Browser opens, you sign in, Clippy captures both access and refresh tokens
2. **Access token expires**: Clippy automatically uses the refresh token to get a new access token via API (no browser needed)
3. **Refresh token rotates**: Microsoft issues a new refresh token with each refresh, keeping the chain alive

This is more reliable than session cookies, which Microsoft can invalidate server-side at any time.

### Background Session Keepalive

To keep your session alive indefinitely:

```bash
# Start keepalive (keeps browser session warm)
clippy keepalive

# Or install as macOS LaunchAgent (recommended)
# Create ~/Library/LaunchAgents/com.clippy.keepalive.plist
```

The keepalive:
- Refreshes the browser session every 10 minutes
- Validates tokens against Microsoft's servers
- Sets a `needs-login` marker if session expires (prevents browser spam)
- After manual re-login, automatically resumes

### Token Storage

Tokens are stored in `~/.config/clippy/`:
- `token-cache.json` - Access token, refresh token, expiry
- `storage-state.json` - Browser cookies/session
- `needs-login` - Marker file when re-auth is required
- `keepalive-health.txt` - Last successful keepalive timestamp

---

## Calendar Commands

### View Calendar

```bash
# Today's events
clippy calendar

# Specific day
clippy calendar --day tomorrow
clippy calendar --day monday
clippy calendar --day 2024-02-15

# Week view
clippy calendar --week

# Include details (description, attendees)
clippy calendar --details
```

### Create Events

```bash
# Basic event
clippy create-event "Team Standup" 09:00 09:30

# With options
clippy create-event "Project Review" 14:00 15:00 \
  --day tomorrow \
  --description "Q1 review meeting" \
  --attendees "alice@company.com,bob@company.com" \
  --teams \
  --room "Conference Room A"

# Find an available room automatically
clippy create-event "Workshop" 10:00 12:00 --find-room

# List available rooms
clippy create-event "x" 10:00 11:00 --list-rooms
```

### Recurring Events

```bash
# Daily standup
clippy create-event "Daily Standup" 09:00 09:15 --repeat daily

# Weekly on specific days
clippy create-event "Team Sync" 14:00 15:00 \
  --repeat weekly \
  --days mon,wed,fri

# Monthly, 10 occurrences
clippy create-event "Monthly Review" 10:00 11:00 \
  --repeat monthly \
  --count 10

# Every 2 weeks until a date
clippy create-event "Sprint Planning" 09:00 11:00 \
  --repeat weekly \
  --every 2 \
  --until 2024-12-31
```

### Update Events

```bash
# List today's events to get index
clippy update-event

# Update by index
clippy update-event 1 --title "New Title"
clippy update-event 2 --start 10:00 --end 11:00
clippy update-event 3 --add-attendee "new@company.com"
clippy update-event 4 --teams        # Add Teams meeting
clippy update-event 5 --no-teams     # Remove Teams meeting
```

### Delete/Cancel Events

```bash
# List your events
clippy delete-event

# Delete event (cancels and notifies attendees if any)
clippy delete-event 1

# With cancellation message
clippy delete-event 2 --message "Sorry, need to reschedule"

# Force delete without notification
clippy delete-event 3 --force-delete
```

### Respond to Invitations

```bash
# List events needing response
clippy respond

# Accept/decline/tentative
clippy respond 1 --accept
clippy respond 2 --decline --message "Conflict with another meeting"
clippy respond 3 --tentative
```

### Find Meeting Times

```bash
# Find free slots for yourself
clippy findtime

# Find times when multiple people are free
clippy findtime --attendees "alice@company.com,bob@company.com"

# Specific duration and date range
clippy findtime --duration 60 --days 5
```

---

## Email Commands

### List & Read Email

```bash
# Inbox (default)
clippy mail

# Other folders
clippy mail sent
clippy mail drafts
clippy mail deleted
clippy mail archive

# Pagination
clippy mail -n 20           # Show 20 emails
clippy mail -p 2            # Page 2

# Filters
clippy mail --unread        # Only unread
clippy mail --search "invoice"

# Read an email
clippy mail -r 1            # Read email #1

# Download attachments
clippy mail -d 3            # Download from email #3
clippy mail -d 3 -o ~/Downloads
```

### Send Email

```bash
# Simple email
clippy send \
  --to "recipient@example.com" \
  --subject "Hello" \
  --body "This is the message body"

# Multiple recipients, CC, BCC
clippy send \
  --to "alice@example.com,bob@example.com" \
  --cc "manager@example.com" \
  --bcc "archive@example.com" \
  --subject "Team Update" \
  --body "..."

# With markdown formatting
clippy send \
  --to "user@example.com" \
  --subject "Update" \
  --body "**Bold text** and a [link](https://example.com)" \
  --markdown

# With attachments
clippy send \
  --to "user@example.com" \
  --subject "Report" \
  --body "Please find attached." \
  --attach "report.pdf,data.xlsx"
```

### Reply & Forward

```bash
# Reply to an email
clippy mail --reply 1 --message "Thanks for your email!"

# Reply all
clippy mail --reply-all 1 --message "Thanks everyone!"

# Reply with markdown
clippy mail --reply 1 --message "**Got it!** Will do." --markdown

# Forward an email
clippy mail --forward 1 --to-addr "colleague@example.com"
clippy mail --forward 1 --to-addr "a@example.com,b@example.com" --message "FYI"
```

### Email Actions

```bash
# Mark as read/unread
clippy mail --mark-read 1
clippy mail --mark-unread 2

# Flag emails
clippy mail --flag 1
clippy mail --unflag 2
clippy mail --complete 3    # Mark flag as complete

# Move to folder
clippy mail --move 1 --to archive
clippy mail --move 2 --to deleted
clippy mail --move 3 --to "My Custom Folder"
```

### Manage Drafts

```bash
# List drafts
clippy drafts

# Read a draft
clippy drafts -r 1

# Create a draft
clippy drafts --create \
  --to "recipient@example.com" \
  --subject "Draft Email" \
  --body "Work in progress..."

# Create with attachment
clippy drafts --create \
  --to "user@example.com" \
  --subject "Report" \
  --body "See attached" \
  --attach "report.pdf"

# Edit a draft
clippy drafts --edit 1 --body "Updated content"
clippy drafts --edit 1 --subject "New Subject"

# Send a draft
clippy drafts --send 1

# Delete a draft
clippy drafts --delete 1
```

### Manage Folders

```bash
# List all folders
clippy folders

# Create a folder
clippy folders --create "Projects"

# Rename a folder
clippy folders --rename "Projects" --new-name "Active Projects"

# Delete a folder
clippy folders --delete "Old Folder"
```

---

## People & Room Search

```bash
# Search for people
clippy find "john"

# Search for rooms
clippy find "conference" --rooms
```

---

## Global Options

All commands support:

```bash
--gcc               # Use Office 365 US Government (GCC) endpoints
--json              # Output as JSON (for scripting)
--token <token>     # Use a specific token
-i, --interactive   # Force interactive browser login
--no-headless       # Show the browser window during login (don't run headless)
```

---

## Examples

### Morning Routine Script

```bash
#!/bin/bash
echo "=== Today's Calendar ==="
clippy calendar

echo -e "\n=== Unread Emails ==="
clippy mail --unread -n 5

echo -e "\n=== Pending Invitations ==="
clippy respond
```

### Quick Meeting Setup

```bash
# Find a time when everyone is free and create the meeting
clippy create-event "Project Kickoff" 14:00 15:00 \
  --day tomorrow \
  --attendees "team@company.com" \
  --teams \
  --find-room \
  --description "Initial project planning session"
```

### Email Report with Attachment

```bash
clippy send \
  --to "manager@company.com" \
  --subject "Weekly Report - $(date +%Y-%m-%d)" \
  --body "Please find this week's report attached." \
  --attach "weekly-report.pdf"
```

---

## Requirements

- [Bun](https://bun.sh) runtime
- macOS (for browser-based authentication via Playwright)
- Microsoft 365 / Outlook account

## License

MIT
