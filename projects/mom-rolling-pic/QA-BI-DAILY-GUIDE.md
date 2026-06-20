# QA Bi-Daily

## Overview

QA Bi-Daily manages recurring QA standups for a dynamic number of projects.

- Separate schedule, team, Google Meet, and WhatsApp configuration per project
- Automatic standup row generation
- WhatsApp reminders and summaries through Fonnte
- One scheduler for all configured projects
- Existing standup data is preserved during sheet initialization and generation

## Deployment

```bash
cd projects/mom-rolling-pic
clasp login
clasp push
```

Apps Script ID:

```text
1Rddhr1q5W-E1p-pnzMAlFzoAdjegaWFQ0BKiurKmNBPa8BDUiQDH5drg
```

## Existing Spreadsheet

Do not run `Initialize Config` on an existing configured spreadsheet. It resets the `Config` sheet.

The existing SIPGN and INADigital/Internal blocks remain compatible with the dynamic configuration reader.

## Add Project

1. Open `QA Bi-Daily → Setup → Add Project`.
2. Enter a unique project name.
3. Complete the new project block in `Config`.
4. Fill team members, schedule, Google Meet, WhatsApp Group ID, and Fonnte Token.
5. Run `Initialize Project Sheets`.
6. Run `Setup Auto Triggers` again.

No source-code or trigger-handler changes are required for additional projects.

## Scheduler

`Setup Auto Triggers` removes the legacy Project A/Project B triggers and creates one scheduler that runs every five minutes.

For every configured project, the scheduler:

1. Checks whether today is a configured standup day.
2. Sends the reminder at the configured time minus its reminder offset.
3. Generates today's rows if they do not exist.
4. Sends the summary at the configured summary time.
5. Uses Script Properties to prevent duplicate sends for the same project, date, and action.

Run `Setup Auto Triggers` again after project or schedule changes.

## Manual Operations

- `Generate Today (All Projects)`
- `Generate for Date`
- `Bulk: Next 3 Months`
- `Bulk: Custom Period`
- `Generate One Project Today`
- `Test Reminder`
- `Test Summary`
- `Get WhatsApp Groups`

Project-specific actions prompt for a project name and display the configured names.

## Supported Schedule

The current standup sheet layout supports:

- Monday
- Wednesday
- Friday

Each project may enable any combination of those days through `Standup Days`.

## Version

Version 2.0: renamed to QA Bi-Daily and refactored for dynamic multi-project support.
