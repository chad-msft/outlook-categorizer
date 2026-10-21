# Outlook Categorizer

A Python command-line tool that uses the Decision-1 model through the TypeSafe SDK
to categorize uncategorized emails in your Classic Outlook Inbox.

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

- Processes emails in the default Inbox only, not subfolders.
- Skips emails that already have a category.
- Assigns and saves one configured category per email.
- Prints the category, model confidence, and a summary of results.

**Important:** This modifies emails directly; there is no dry-run mode. The
sender's email address, subject, and first 1,024 characters of the body are sent to
the configured model endpoint. Use only an endpoint approved for your email data.

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

Exit codes: `0` for success, `1` for login, setup, or email errors, and `2` for
invalid command-line arguments.
