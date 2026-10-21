"""Categorize Classic Outlook Inbox emails or Calendar events using a JSON Choice question.

Requires Windows, Classic Outlook, Azure CLI, pywin32, azure-identity, and
typesafe-sdk. Run with --help for arguments and an example. Azure login runs
before mailbox access. Exit codes: 0 for success, 1 for login/setup/item
errors, and 2 for invalid command-line arguments.
"""

import argparse
import json
import logging
import re
import shutil
import subprocess
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlsplit

import pywintypes
import win32api
import win32com.client
import winreg
from azure.core.exceptions import AzureError
from azure.identity import AzureCliCredential, get_bearer_token_provider
from typesafe_sdk import Choice, TypeSafeClient, TypeSafeError


def load_category_choice(choices_path: Path) -> Choice:
    try:
        config = json.loads(choices_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Could not read valid Choice JSON from {choices_path.resolve()}.") from exc
    if not isinstance(config, dict) or set(config) != {"category"}:
        raise ValueError("The choices file must contain exactly one 'category' Choice question.")
    question = Choice.model_validate(config["category"])
    if not question.criteria:
        raise ValueError("The category Choice question must have at least one criterion.")
    for name, description in question.criteria.items():
        if not name.strip() or not isinstance(description, str) or not description.strip():
            raise ValueError(f"Category {name!r} needs a non-empty name and text description.")
    return question


def outlook_sender_smtp_address(email) -> str:
    """Resolve internal Exchange senders to SMTP, not legacy Exchange names."""
    if email.SenderEmailType != "EX":
        address = email.SenderEmailAddress
    else:
        sender = email.Sender
        if sender is None:
            raise ValueError("The email has no resolvable Exchange sender.")
        if sender.AddressEntryUserType in (0, 5):
            exchange_user = sender.GetExchangeUser()
            address = exchange_user.PrimarySmtpAddress if exchange_user is not None else ""
        elif sender.AddressEntryUserType == 1:
            exchange_list = sender.GetExchangeDistributionList()
            address = exchange_list.PrimarySmtpAddress if exchange_list is not None else ""
        else:
            address = sender.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
            )
    if not isinstance(address, str) or "@" not in address:
        raise ValueError("Could not resolve the sender to an SMTP email address.")
    return address.strip()


def log_outlook_item_error(item, index: int, stage: str, is_calendar: bool) -> None:
    details = {}
    fields = (
        ("EntryID", "Start", "End", "Organizer", "Subject")
        if is_calendar
        else ("EntryID", "ReceivedTime", "SenderEmailAddress", "Subject")
    )
    for field in fields:
        try:
            details[field] = getattr(item, field)
        except (pywintypes.com_error, AttributeError) as metadata_error:
            details[field] = f"<unavailable: {metadata_error}>"
    logging.getLogger("outlook_categorization").exception(
        "%s %d failed while %s; details=%r",
        "Calendar event" if is_calendar else "Email",
        index,
        stage,
        details,
    )


def login_to_azure(tenant_id: str) -> None:
    executable = shutil.which("az")
    if executable is None:
        raise RuntimeError("Azure CLI was not found. Install it and ensure 'az' is on PATH.")
    subprocess.run(
        [executable, "login", "--tenant", tenant_id, "--output", "none"],
        check=True,
    )


def categorize_outlook_items(
    start_at: datetime,
    end_at: datetime,
    category_question: Choice,
    base_url: str,
    tenant_id: str,
    process_calendar: bool = False,
) -> int:
    """Process Inbox emails or overlapping Calendar events; return the error count."""
    if start_at > end_at:
        raise ValueError("start_date must be on or before end_date.")
    end_exclusive = end_at + timedelta(days=1)
    category_criteria = category_question.criteria
    locale_id = win32api.GetUserDefaultLCID()
    date_shortdate, time_noseconds = 1, 2

    def outlook_date_literal(value: datetime) -> str:
        date_text = win32api.GetDateFormat(locale_id, date_shortdate, value)
        time_text = win32api.GetTimeFormat(locale_id, time_noseconds, value)
        return f"{date_text} {time_text}"

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\International") as key:
            category_separator, _ = winreg.QueryValueEx(key, "sList")
    except OSError as exc:
        raise RuntimeError("Could not read Outlook's Windows category list separator.") from exc
    if not isinstance(category_separator, str) or not category_separator.strip():
        raise ValueError("Windows has an invalid category list separator.")
    for name in category_criteria:
        if category_separator in name:
            raise ValueError(f"Category {name!r} contains Outlook's category list separator.")

    total = skipped = requests_sent = processed = updated = errors = 0
    def print_summary() -> None:
        item_plural = "Calendar events" if process_calendar else "Emails"
        print("\nCategorization summary")
        print(f"Total {item_plural.lower()} in timeframe: {total}")
        print(
            f"{item_plural} processed by model: "
            f"{processed} (requests attempted: {requests_sent})"
        )
        print(f"Skipped because already categorized: {skipped}")
        print(f"Categories saved: {updated}")
        print(f"{'Calendar event' if process_calendar else 'Email'} errors: {errors}")

    folder_description = "Calendar events" if process_calendar else "Inbox messages"
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folders = [(9, 26, True)] if process_calendar else [(6, 43, False)]

        # Snapshot uncategorized matches before saving any category changes.
        pending_items = []
        for folder_id, item_class, is_calendar in folders:
            folder_description = "Calendar events" if is_calendar else "Inbox messages"
            folder = namespace.GetDefaultFolder(folder_id)
            if is_calendar:
                items = folder.Items
                items.Sort("[Start]")
                items.IncludeRecurrences = True
                # Jet compares Start/End in local time; include events spanning either boundary.
                calendar_query = (
                    f"[Start] < '{outlook_date_literal(end_exclusive)}' AND "
                    f"[End] > '{outlook_date_literal(start_at)}'"
                )
                items = items.Restrict(calendar_query)
            else:
                # DASL uses UTC boundaries and Windows-localized date/time literals.
                received_property = '"urn:schemas:httpmail:datereceived"'
                email_query = (
                    f"@SQL={received_property} >= "
                    f"'{outlook_date_literal(start_at.astimezone(timezone.utc))}' AND "
                    f"{received_property} < "
                    f"'{outlook_date_literal(end_exclusive.astimezone(timezone.utc))}'"
                )
                items = folder.Items.Restrict(email_query)
                items.Sort("[ReceivedTime]", True)

            item = items.GetFirst()
            while item is not None:
                if item.Class == item_class:
                    # Bound enumeration even for recurring series without an end date.
                    if is_calendar and item.Start.replace(tzinfo=None) >= end_exclusive:
                        break
                    total += 1
                    try:
                        if (item.Categories or "").strip():
                            skipped += 1
                        else:
                            pending_items.append((total, item, is_calendar))
                    except (pywintypes.com_error, AttributeError):
                        errors += 1
                        log_outlook_item_error(
                            item, total, "checking existing categories", is_calendar
                        )
                item = items.GetNext()
        if not pending_items:
            print_summary()
            return errors
    except pywintypes.com_error as exc:
        raise RuntimeError(
            f"Could not read {folder_description}. Open Classic Outlook for Windows "
            "and sign in to your mail profile, then rerun this script."
        ) from exc

    with AzureCliCredential(tenant_id=tenant_id) as credential:
        token_provider = get_bearer_token_provider(credential, "https://ai.azure.com/.default")
        with TypeSafeClient(
            api_key=token_provider(),
            base_url=base_url,
            model="Decision-1",
            timeout=60,
        ) as client:
            for index, item, is_calendar in pending_items:
                item_name = "Calendar event" if is_calendar else "Email"
                stage = "checking existing categories"
                try:
                    if (item.Categories or "").strip():
                        skipped += 1
                        continue
                    stage = "reading the classification payload"
                    subject = item.Subject or ""
                    body_preview = (item.Body or "")[:1024]
                    if is_calendar:
                        state = (
                            f"Organizer: {item.Organizer or ''}\nSubject: {subject}\n"
                            f"Start: {item.Start}\nEnd: {item.End}\n"
                            f"All-day: {item.AllDayEvent}\nLocation: {item.Location or ''}\n"
                            f"Body: {body_preview}"
                        )
                    else:
                        sender_email = outlook_sender_smtp_address(item)
                        state = f"Sender: {sender_email}\nSubject: {subject}\nBody: {body_preview}"
                    stage = "requesting model categorization"
                    requests_sent += 1
                    response = client.system_one(
                        state=state,
                        questions={"category": category_question},
                    )
                    stage = "validating the model response"
                    answer = response.choices["category"]
                    selected = answer.choice
                    if selected not in category_criteria:
                        raise ValueError(f"Model returned an unknown configured category: {selected!r}")
                    if not 0.0 <= answer.confidence <= 1.0:
                        raise ValueError(
                            f"Invalid category confidence for {item_name.lower()} {index}: "
                            f"{answer.confidence!r}"
                        )
                    processed += 1

                    stage = "saving the selected category"
                    if (item.Categories or "").strip():
                        status = "Categorized during model request; unchanged"
                    else:
                        item.Categories = selected
                        item.Save()
                        updated += 1
                        status = "Category assigned and saved"
                except (pywintypes.com_error, TypeSafeError, ValueError, KeyError, AttributeError):
                    errors += 1
                    log_outlook_item_error(item, index, stage, is_calendar)
                    continue
                display_subject = subject or "(no subject)"
                if process_calendar:
                    display_subject = f"{item_name}: {display_subject}"
                print(
                    f"{index:2}. {display_subject}\n"
                    f"    Category: {selected} | Confidence: {answer.confidence:.1%} | {status}"
                )

    print_summary()
    return errors


def date_argument(value: str) -> datetime:
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Use a valid date in YYYY-MM-DD format.") from exc


def tenant_argument(value: str) -> str:
    if re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9.-]*", value) is None:
        raise argparse.ArgumentTypeError("Use a tenant ID or tenant domain, without spaces or shell characters.")
    return value


def base_url_argument(value: str) -> str:
    try:
        parsed = urlsplit(value)
        valid = (
            parsed.scheme == "https"
            and parsed.hostname
            and not parsed.username
            and not parsed.password
            and not parsed.query
            and not parsed.fragment
            and not any(character.isspace() for character in value)
        )
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Use a valid HTTPS base URL.") from exc
    if not valid:
        raise argparse.ArgumentTypeError("Use an HTTPS base URL without credentials, query, or fragment.")
    return value


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Sign in to Azure, then categorize uncategorized Classic Outlook Inbox emails "
            "or Calendar events using a JSON Choice question. "
            "Running this script saves item categories."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Requires Windows, Classic Outlook signed in, and Azure CLI on PATH.\n"
            "Python dependencies (from the repository root):\n"
            "  python -m pip install -r .\\requirements.txt\n\n"
            "Example from the repository root:\n"
            "  python .\\outlook-categorizer.py "
            '--base-url "https://YOUR-RESOURCE.services.ai.azure.com/providers/microsoft" '
            '--choices-path ".\\choices.json" --start-date 2026-07-01 '
            '--end-date 2026-07-31 --tenant-id "YOUR-TENANT-ID"\n\n'
            "Runs az login --tenant TENANT_ID before accessing Outlook.\n"
            "Dates are inclusive in the Windows local time zone. Existing categories are skipped.\n"
            "Use --process-calendar to categorize overlapping Calendar events instead of Inbox emails.\n"
            "Exit codes: 0 success; 1 login, setup, or per-item errors; 2 invalid arguments."
        ),
    )
    parser.add_argument("--base-url", required=True, type=base_url_argument, help="Decisions HTTPS base URL.")
    parser.add_argument(
        "--choices-path",
        required=True,
        type=Path,
        help="Choice JSON file; relative paths are resolved from the current working directory.",
    )
    parser.add_argument("--start-date", required=True, type=date_argument, help="First local date, YYYY-MM-DD.")
    parser.add_argument("--tenant-id", required=True, type=tenant_argument, help="Azure tenant ID or tenant domain.")
    parser.add_argument("--end-date", required=True, type=date_argument, help="Last local date, YYYY-MM-DD.")
    parser.add_argument(
        "--process-calendar",
        action="store_true",
        help=(
            "Categorize default Calendar events instead of Inbox emails, "
            "overlapping the inclusive local date range."
        ),
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.start_date > args.end_date:
        parser.error("--start-date must be on or before --end-date.")
    if args.end_date.date() == datetime.max.date():
        parser.error("--end-date must be earlier than 9999-12-31.")
    logging.basicConfig(level=logging.ERROR, format="%(asctime)s %(levelname)s %(message)s")
    try:
        category_question = load_category_choice(args.choices_path.expanduser().resolve())
        login_to_azure(args.tenant_id)
        errors = categorize_outlook_items(
            start_at=args.start_date,
            end_at=args.end_date,
            category_question=category_question,
            base_url=args.base_url,
            tenant_id=args.tenant_id,
            process_calendar=args.process_calendar,
        )
    except (AzureError, OSError, RuntimeError, TypeSafeError, ValueError, subprocess.CalledProcessError):
        logging.getLogger("outlook_categorization").exception("Outlook categorization failed.")
        return 1
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
