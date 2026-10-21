# Outlook Categorizer

A Python command-line tool that uses the Decision-1 model through the TypeSafe SDK
to categorize uncategorized emails in your Classic Outlook Inbox or events in
your default Calendar.

## Requirements

- Windows with Classic Outlook installed and signed in to your mail profile
  (new Outlook is not supported).
- Python 3.10 or later.
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli-windows)
  installed and available on your PATH.
- An Azure account with access to a Decisions endpoint supporting Decision-1.

## Setup

From this project folder, install the dependencies:

```powershell
python -m pip install -r .\requirements.txt
```

## Usage

Open Classic Outlook, then run the following in PowerShell. Replace the endpoint,
tenant ID, and dates with your own values:

```powershell
python .\outlook-categorizer.py `
  --base-url "https://YOUR-RESOURCE.services.ai.azure.com/providers/microsoft" `
  --tenant-id "YOUR-TENANT-ID" `
  --choices-path ".\choices.json" `
  --start-date 2026-10-01 `
  --end-date 2026-10-21
```

The script runs `az login` for the specified tenant before accessing Outlook.
Both dates are inclusive and use your Windows local time zone.

- Processes default Inbox emails without `--process-calendar`, or default Calendar
  events with it, never both. Subfolders are not scanned.
- Skips items that already have a category.
- Assigns and saves one configured category per item.
- Prints the category, model confidence, and a summary of results.

### Calendar events

Add `--process-calendar` to the command above to process the default Calendar
**instead of Inbox emails**. Without this flag, only Inbox emails are processed.

Calendar events are selected when they overlap any part of the inclusive
`--start-date` to `--end-date` range, based on their start/end times rather than
when they were received or created. This includes all-day and multi-day events
that began before `--start-date` but are still underway. An event ending exactly
at midnight at the start of the range, or starting at midnight after `--end-date`,
is excluded.

Recurring occurrences within the range are processed individually, not the
entire series. Shared calendars and Calendar subfolders are not scanned.
Existing event categories are preserved.

**Important:** This modifies items directly in the selected processing mode;
there is no dry-run mode. For emails, the
sender's email address, subject, and first 1,024 characters of the body are sent to
the configured model endpoint. For calendar events, the organizer name, subject,
start/end times, all-day status, location, and first 1,024 body characters are sent.
Use only an endpoint approved for your email and calendar data.

## Categories

Edit [choices.json](choices.json) to customize the categories. Keep the single
top-level `category` question with `type` set to `choice`. Update `instructions`
and the `criteria` mapping, where each key is an Outlook category name and each
value describes when to choose it. Names and descriptions must not be blank.
Category names must not contain your Windows list separator.

For all command-line options:

```powershell
python .\outlook-categorizer.py --help
```

Exit codes: `0` for success, `1` for login, setup, or item errors, and `2` for
invalid command-line arguments.

## Tests

Run the mocked tests without accessing Outlook or Azure:

```powershell
python -B -m unittest discover -v
```
