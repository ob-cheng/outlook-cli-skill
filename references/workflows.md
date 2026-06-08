# Common Workflows

Patterns for common email, calendar, and task management scenarios.

## Email Workflows

### Inbox Triage

Find and process unread emails:

```bash
# 1. Get all unread emails
python "${SKILL_DIR}/outlook.py" search --unread --json

# 2. Read specific email for details
python "${SKILL_DIR}/outlook.py" read <message-id> --json

# 3. Reply or forward as needed
python "${SKILL_DIR}/outlook.py" reply <message-id> --body "Thanks, I'll handle this"
```

### Find Emails from Someone

Search for emails from a specific person:

```bash
# Last 14 days from specific email
python "${SKILL_DIR}/outlook.py" search --filter-email boss@company.com --days 14 --json

# From a domain
python "${SKILL_DIR}/outlook.py" search --filter-domain vendor.com --days 30 --json

# Combine with keyword
python "${SKILL_DIR}/outlook.py" search --filter-email client@co.com --keyword "contract" --json
```

### Project Email Archive

Export all project-related emails:

```bash
# To markdown files (for Obsidian)
python "${SKILL_DIR}/outlook.py" export --output ./project-archive \
  --filter-email client@company.com \
  --keyword "ProjectX" \
  --days 90

# To JSON for processing
python "${SKILL_DIR}/outlook.py" export --output ./data \
  --filter-domain vendor.com \
  --format json --batch

# Direct to stdout for AI consumption
python "${SKILL_DIR}/outlook.py" export --output . --stdout \
  --filter-email stakeholder@co.com \
  --days 30
```

### Incremental Export

Set up recurring exports that only capture new emails:

```bash
# First run - exports all matching emails, saves state
python "${SKILL_DIR}/outlook.py" export --output ./daily-export --incremental --days 30

# Subsequent runs - only exports emails since last run
python "${SKILL_DIR}/outlook.py" export --output ./daily-export --incremental
```

### Send with Attachments

Compose email with multiple attachments:

```bash
python "${SKILL_DIR}/outlook.py" send \
  --to "recipient@company.com" \
  --cc "manager@company.com" \
  --subject "Q2 Report" \
  --body "Please find the Q2 reports attached." \
  --attach ./reports/q2-summary.pdf \
  --attach ./reports/q2-data.xlsx
```

### Draft Review Flow

Create drafts for review before sending:

```bash
# Create draft
python "${SKILL_DIR}/outlook.py" send \
  --to "client@company.com" \
  --subject "Proposal" \
  --body "Here's our proposal..." \
  --draft

# Review in Outlook, then send manually
# Or create reply draft
python "${SKILL_DIR}/outlook.py" reply <message-id> --body "Response text" --draft
```

---

## Calendar Workflows

### Daily Schedule Check

View today's and upcoming meetings:

```bash
# Today's events
python "${SKILL_DIR}/outlook.py" cal list --json

# This week
python "${SKILL_DIR}/outlook.py" cal list --end 2026-06-14 --json

# Filter by location
python "${SKILL_DIR}/outlook.py" cal list --location "Room A" --json
```

### Schedule Team Meeting

Create meeting with required and optional attendees:

```bash
python "${SKILL_DIR}/outlook.py" cal create \
  --subject "Project Kickoff" \
  --start "2026-06-10 14:00" \
  --end "2026-06-10 15:30" \
  --location "Conference Room B" \
  --body "Kickoff meeting for Q3 initiative" \
  --required "alice@co.com,bob@co.com,carol@co.com" \
  --optional "dan@co.com" \
  --reminder 30
```

### Find Free Time

Check calendar for conflicts:

```bash
# Check specific date range
python "${SKILL_DIR}/outlook.py" cal list --start 2026-06-10 --end 2026-06-10 --json

# Filter out all-day events
python "${SKILL_DIR}/outlook.py" cal list --start 2026-06-10 --end 2026-06-12 --json | jq '.events | map(select(.is_all_day == false))'
```

### Recurring Meeting Audit

Find all recurring meetings:

```bash
python "${SKILL_DIR}/outlook.py" cal list --recurring --start 2026-06-01 --end 2026-06-30 --json
```

---

## Task Workflows

### Task Dashboard

View pending work:

```bash
# All incomplete tasks
python "${SKILL_DIR}/outlook.py" tasks list --json

# High priority only
python "${SKILL_DIR}/outlook.py" tasks list --priority high --json

# Due this week
python "${SKILL_DIR}/outlook.py" tasks list --due-before 2026-06-14 --json

# Overdue tasks
python "${SKILL_DIR}/outlook.py" tasks list --due-before 2026-06-08 --json
```

### Create Task from Email

After reading an email that requires action:

```bash
# 1. Read the email
python "${SKILL_DIR}/outlook.py" read <message-id> --json

# 2. Create a follow-up task
python "${SKILL_DIR}/outlook.py" tasks create \
  --subject "Follow up: [email subject]" \
  --due 2026-06-15 \
  --priority high \
  --body "Reference email from [sender]"
```

### Task Completion Flow

```bash
# 1. List pending tasks
python "${SKILL_DIR}/outlook.py" tasks list --json

# 2. Work on task, then mark complete
python "${SKILL_DIR}/outlook.py" tasks complete <task-id>

# 3. Verify completion
python "${SKILL_DIR}/outlook.py" tasks list --all --json
```

### Weekly Review

```bash
# Tasks completed this week
python "${SKILL_DIR}/outlook.py" tasks list --status completed --json

# Tasks due next week
python "${SKILL_DIR}/outlook.py" tasks list --due-after 2026-06-08 --due-before 2026-06-15 --json

# All tasks by category
python "${SKILL_DIR}/outlook.py" tasks list --category "Project A" --json
```

---

## AI Agent Workflows

### Email Intelligence Pipeline

Export emails for AI analysis:

```bash
# Get recent emails as JSON for processing
python "${SKILL_DIR}/outlook.py" export --output . --stdout --days 7 > emails.json

# Filter specific conversations
python "${SKILL_DIR}/outlook.py" export --output . --stdout \
  --filter-email important@client.com \
  --days 30 > client-emails.json
```

### Automated Triage

```bash
# 1. Get unread count
python "${SKILL_DIR}/outlook.py" search --unread --json | jq '.count'

# 2. Get summaries for triage
python "${SKILL_DIR}/outlook.py" search --unread --json | jq '.emails[] | {subject, sender, date}'

# 3. Read specific email for AI processing
python "${SKILL_DIR}/outlook.py" read <message-id> --json | jq '.emails[0].body'
```

### Calendar + Task Integration

Prepare for meetings by checking related tasks:

```bash
# 1. Get today's meetings
python "${SKILL_DIR}/outlook.py" cal list --json

# 2. For each meeting, check related tasks
python "${SKILL_DIR}/outlook.py" tasks list --keyword "Project X" --json
```

---

## Multi-Step Patterns

### Email Thread Analysis

```bash
# 1. Search for thread
python "${SKILL_DIR}/outlook.py" search --keyword "Q2 Planning" --days 30 --json

# 2. Export full thread
python "${SKILL_DIR}/outlook.py" export --output . --stdout --keyword "Q2 Planning" --days 30

# 3. Read specific messages for detail
python "${SKILL_DIR}/outlook.py" read <id1> <id2> <id3> --json
```

### Meeting Prep Checklist

```bash
# 1. Get meeting details
python "${SKILL_DIR}/outlook.py" cal read <event-id> --json

# 2. Find related emails from attendees
python "${SKILL_DIR}/outlook.py" search --filter-email attendee@co.com --days 14 --json

# 3. Check related tasks
python "${SKILL_DIR}/outlook.py" tasks list --keyword "meeting topic" --json

# 4. Create prep task if needed
python "${SKILL_DIR}/outlook.py" tasks create --subject "Prep for [meeting]" --due 2026-06-09
```
