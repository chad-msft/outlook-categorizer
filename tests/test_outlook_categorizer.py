import importlib.util
import io
import operator
import re
import unittest
from contextlib import ExitStack, redirect_stdout
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock, patch


spec = importlib.util.spec_from_file_location(
    "outlook_categorizer", Path(__file__).with_name("outlook-categorizer.py")
)
if spec is None or spec.loader is None:
    raise RuntimeError("Could not load the Outlook categorizer for testing.")
categorizer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(categorizer)


class OutlookItems:
    def __init__(self, items=(), occurrences=()):
        self.items = list(items)
        self.occurrences = list(occurrences)
        self.operations = []
        self.include_recurrences = False
        self.sort_field = None
        self.descending = False
        self.index = 0

    @property
    def IncludeRecurrences(self):
        return self.include_recurrences

    @IncludeRecurrences.setter
    def IncludeRecurrences(self, value):
        self.operations.append(("recurrences", value))
        self.include_recurrences = value

    @property
    def Count(self):
        raise AssertionError("Recurring collections must not be enumerated using Count.")

    def Sort(self, field, descending=False):
        self.operations.append(("sort", field, descending))
        self.sort_field = field.strip("[]")
        self.descending = descending
        self.items.sort(key=lambda item: getattr(item, self.sort_field), reverse=descending)

    def Restrict(self, query):
        self.operations.append(("restrict", query))
        candidates = self.items + (self.occurrences if self.include_recurrences else [])
        comparisons = {
            "<": operator.lt,
            ">": operator.gt,
            ">=": operator.ge,
            "<=": operator.le,
        }
        if query.startswith("@SQL="):
            clauses = re.findall(
                r'"urn:schemas:httpmail:datereceived" ([<>=]+) \'([^\']+)\'', query
            )
            checks = [("ReceivedTime", comparison, value) for comparison, value in clauses]
        else:
            checks = re.findall(r"\[(Start|End)\] ([<>=]+) '([^']+)'", query)
        if len(checks) != 2:
            raise AssertionError(f"Unexpected Outlook date filter: {query}")

        selected = [
            item
            for item in candidates
            if all(
                comparisons[comparison](
                    getattr(item, field).replace(tzinfo=None),
                    datetime.strptime(value, "%m/%d/%Y %H:%M"),
                )
                for field, comparison, value in checks
            )
        ]
        if self.sort_field:
            selected.sort(
                key=lambda item: getattr(item, self.sort_field), reverse=self.descending
            )
        return OutlookItems(selected)

    def GetFirst(self):
        self.index = 0
        return self.items[0] if self.items else None

    def GetNext(self):
        self.index += 1
        return self.items[self.index] if self.index < len(self.items) else None


def email(subject="Email", categories=""):
    return SimpleNamespace(
        Class=43,
        EntryID=subject,
        ReceivedTime=datetime(2026, 10, 21, 12, tzinfo=timezone.utc),
        SenderEmailType="SMTP",
        SenderEmailAddress="sender@example.com",
        Subject=subject,
        Body="Email body",
        Categories=categories,
        Save=Mock(),
    )


def event(subject="Event", start=None, end=None, categories="", all_day=False):
    return SimpleNamespace(
        Class=26,
        EntryID=subject,
        Start=start or datetime(2026, 10, 21, 12),
        End=end or datetime(2026, 10, 21, 13),
        Organizer="Meeting organizer",
        Location="Meeting room",
        AllDayEvent=all_day,
        Subject=subject,
        Body="Event body",
        Categories=categories,
        Save=Mock(),
    )


class OutlookCategorizerTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.choice = categorizer.Choice.model_validate(
            {
                "type": "choice",
                "instructions": "Choose a category.",
                "criteria": {"Important": "Important content."},
            }
        )
        self.inbox = SimpleNamespace(Items=OutlookItems())
        self.calendar = SimpleNamespace(Items=OutlookItems())
        self.namespace = Mock()
        self.namespace.GetDefaultFolder.side_effect = {
            6: self.inbox,
            9: self.calendar,
        }.__getitem__
        outlook = Mock()
        outlook.GetNamespace.return_value = self.namespace
        self.stack.enter_context(
            patch.object(categorizer.win32com.client, "Dispatch", return_value=outlook)
        )
        self.stack.enter_context(
            patch.object(categorizer.win32api, "GetUserDefaultLCID", return_value=1033)
        )
        self.stack.enter_context(
            patch.object(
                categorizer.win32api,
                "GetDateFormat",
                side_effect=lambda _locale, _flags, value: value.strftime("%m/%d/%Y"),
            )
        )
        self.stack.enter_context(
            patch.object(
                categorizer.win32api,
                "GetTimeFormat",
                side_effect=lambda _locale, _flags, value: value.strftime("%H:%M"),
            )
        )
        self.stack.enter_context(
            patch.object(categorizer.winreg, "OpenKey", return_value=MagicMock())
        )
        self.stack.enter_context(
            patch.object(categorizer.winreg, "QueryValueEx", return_value=(",", 1))
        )
        self.stack.enter_context(patch.object(categorizer, "AzureCliCredential"))
        self.stack.enter_context(
            patch.object(
                categorizer, "get_bearer_token_provider", return_value=lambda: "test-token"
            )
        )
        self.client_class = self.stack.enter_context(patch.object(categorizer, "TypeSafeClient"))
        self.client = self.client_class.return_value.__enter__.return_value
        self.client.system_one.return_value = SimpleNamespace(
            choices={"category": SimpleNamespace(choice="Important", confidence=0.9)}
        )
        self.output = io.StringIO()
        self.stack.enter_context(redirect_stdout(self.output))
        self.arguments = [
            "--base-url", "https://example.com/providers/microsoft",
            "--tenant-id", "example.com",
            "--choices-path", "choices.json",
            "--start-date", "2026-10-21",
            "--end-date", "2026-10-21",
        ]

    def run_categorization(self, process_calendar=False, start_at=None, end_at=None):
        return categorizer.categorize_outlook_items(
            start_at=start_at or datetime(2026, 10, 21),
            end_at=end_at or datetime(2026, 10, 21),
            category_question=self.choice,
            base_url="https://example.com/providers/microsoft",
            tenant_id="example.com",
            process_calendar=process_calendar,
        )

    def test_cli_calendar_flag(self):
        parser = categorizer.build_parser()
        self.assertFalse(parser.parse_args(self.arguments).process_calendar)
        self.assertTrue(
            parser.parse_args(self.arguments + ["--process-calendar"]).process_calendar
        )
        self.assertIn("--process-calendar", parser.format_help())
        self.assertIn(
            "Calendar events instead of Inbox emails", " ".join(parser.format_help().split())
        )

    def test_main_forwards_calendar_flag_and_errors(self):
        with (
            patch.object(categorizer, "load_category_choice", return_value=self.choice),
            patch.object(categorizer, "login_to_azure") as login,
            patch.object(categorizer, "categorize_outlook_items") as process,
        ):
            for enabled in (False, True):
                for errors in (0, 1):
                    with self.subTest(enabled=enabled, errors=errors):
                        process.return_value = errors
                        arguments = self.arguments + (["--process-calendar"] if enabled else [])
                        self.assertEqual(categorizer.main(arguments), errors)
                        self.assertEqual(process.call_args.kwargs["process_calendar"], enabled)
                        login.assert_called_with("example.com")

    def test_default_processes_only_uncategorized_inbox_emails(self):
        message = email()
        categorized = email("Already categorized", "Personal")
        invitation = email("Meeting invitation")
        invitation.Class = 53
        outside = email("Outside dates")
        outside.ReceivedTime = datetime(2026, 10, 23, tzinfo=timezone.utc)
        self.inbox.Items = OutlookItems([message, categorized, invitation, outside])
        appointment = event()
        self.calendar.Items = OutlookItems([appointment])
        self.assertEqual(self.run_categorization(), 0)
        self.namespace.GetDefaultFolder.assert_called_once_with(6)
        self.assertEqual(message.Categories, "Important")
        message.Save.assert_called_once_with()
        categorized.Save.assert_not_called()
        invitation.Save.assert_not_called()
        outside.Save.assert_not_called()
        appointment.Save.assert_not_called()
        self.client.system_one.assert_called_once_with(
            state="Sender: sender@example.com\nSubject: Email\nBody: Email body",
            questions={"category": self.choice},
        )
        self.assertIn("Total emails in timeframe: 2", self.output.getvalue())
        self.assertIn("Email errors: 0", self.output.getvalue())

    def test_calendar_flag_processes_only_calendar_events_and_event_payload(self):
        message = email()
        appointment = event()
        appointment.Body = "x" * 1100
        self.inbox.Items = OutlookItems([message])
        self.calendar.Items = OutlookItems([appointment])
        self.assertEqual(self.run_categorization(process_calendar=True), 0)
        self.namespace.GetDefaultFolder.assert_called_once_with(9)
        self.assertEqual(message.Categories, "")
        message.Save.assert_not_called()
        self.assertEqual(appointment.Categories, "Important")
        appointment.Save.assert_called_once_with()
        self.client.system_one.assert_called_once()
        state = self.client.system_one.call_args.kwargs["state"]
        for value in ("Meeting organizer", "Event", "2026-10-21", "Meeting room"):
            self.assertIn(value, state)
        self.assertEqual(state.split("Body: ", 1)[1], "x" * 1024)
        self.assertIn("Total calendar events in timeframe: 1", self.output.getvalue())
        self.assertIn("Calendar events processed by model: 1", self.output.getvalue())
        self.assertIn("Calendar event errors: 0", self.output.getvalue())
        self.assertNotIn("Total emails", self.output.getvalue())

    def test_calendar_overlap_boundaries_all_day_and_multiday_events(self):
        appointments = [
            event("Earlier", datetime(2026, 10, 19), datetime(2026, 10, 20)),
            event("Ends at start", datetime(2026, 10, 20), datetime(2026, 10, 21)),
            event("Spans range", datetime(2026, 10, 19), datetime(2026, 10, 23)),
            event("Overnight", datetime(2026, 10, 20, 23), datetime(2026, 10, 21, 1)),
            event("All day", datetime(2026, 10, 21), datetime(2026, 10, 22), all_day=True),
            event("Last minute", datetime(2026, 10, 21, 23, 59), datetime(2026, 10, 22, 1)),
            event("Starts at end", datetime(2026, 10, 22), datetime(2026, 10, 22, 1)),
        ]
        self.calendar.Items = OutlookItems(appointments)
        self.assertEqual(self.run_categorization(process_calendar=True), 0)
        for appointment in appointments:
            expected = appointment.Subject in {"Spans range", "Overnight", "All day", "Last minute"}
            with self.subTest(appointment=appointment.Subject):
                self.assertEqual(appointment.Save.call_count, int(expected))
        self.assertEqual(
            self.calendar.Items.operations[-1],
            (
                "restrict",
                "[Start] < '10/22/2026 00:00' AND [End] > '10/21/2026 00:00'",
            ),
        )

    def test_recurring_occurrences_are_expanded_sorted_and_saved_individually(self):
        occurrences = [
            event("Later occurrence", datetime(2026, 10, 28, 12), datetime(2026, 10, 28, 13)),
            event("In-range occurrence"),
            event("Earlier occurrence", datetime(2026, 10, 14, 12), datetime(2026, 10, 14, 13)),
        ]
        self.calendar.Items = OutlookItems(occurrences=occurrences)
        self.assertEqual(self.run_categorization(process_calendar=True), 0)
        self.assertEqual(
            self.calendar.Items.operations[:2],
            [("sort", "[Start]", False), ("recurrences", True)],
        )
        self.assertEqual(self.calendar.Items.operations[2][0], "restrict")
        self.assertEqual([item.Save.call_count for item in occurrences], [0, 1, 0])

    def test_calendar_range_includes_the_entire_end_date_in_local_time(self):
        local_timezone = timezone(timedelta(hours=-7))
        appointments = [
            event(
                "Last day",
                datetime(2026, 10, 23, 23, 30, tzinfo=local_timezone),
                datetime(2026, 10, 24, 0, 30, tzinfo=local_timezone),
            ),
            event(
                "After last day",
                datetime(2026, 10, 24, tzinfo=local_timezone),
                datetime(2026, 10, 24, 1, tzinfo=local_timezone),
            ),
        ]
        self.calendar.Items = OutlookItems(appointments)
        self.assertEqual(
            self.run_categorization(
                process_calendar=True,
                start_at=datetime(2026, 10, 21),
                end_at=datetime(2026, 10, 23),
            ),
            0,
        )
        appointments[0].Save.assert_called_once_with()
        appointments[1].Save.assert_not_called()
        self.assertEqual(
            self.calendar.Items.operations[-1],
            (
                "restrict",
                "[Start] < '10/24/2026 00:00' AND [End] > '10/21/2026 00:00'",
            ),
        )

    def test_calendar_enumeration_stops_at_end_bound_for_open_ended_series(self):
        current = event()
        future = event("Future", datetime(2026, 10, 22), datetime(2026, 10, 22, 1))
        expanded = OutlookItems([current, future])
        expanded.GetNext = Mock(wraps=expanded.GetNext)
        self.calendar.Items.Restrict = Mock(return_value=expanded)
        self.assertEqual(self.run_categorization(process_calendar=True), 0)
        current.Save.assert_called_once_with()
        future.Save.assert_not_called()
        expanded.GetNext.assert_called_once_with()

    def test_existing_calendar_categories_are_skipped(self):
        appointment = event(categories="Personal")
        self.calendar.Items = OutlookItems([appointment])
        self.assertEqual(self.run_categorization(process_calendar=True), 0)
        appointment.Save.assert_not_called()
        self.client_class.assert_not_called()
        self.assertIn("Skipped because already categorized: 1", self.output.getvalue())

    def test_calendar_mode_does_not_require_inbox_access(self):
        appointment = event()
        self.calendar.Items = OutlookItems([appointment])
        self.namespace.GetDefaultFolder.side_effect = {9: self.calendar}.__getitem__
        self.assertEqual(self.run_categorization(process_calendar=True), 0)
        appointment.Save.assert_called_once_with()
        self.namespace.GetDefaultFolder.assert_called_once_with(9)

    def test_empty_calendar_does_not_fall_back_to_inbox(self):
        message = email()
        self.inbox.Items = OutlookItems([message])
        self.assertEqual(self.run_categorization(process_calendar=True), 0)
        self.namespace.GetDefaultFolder.assert_called_once_with(9)
        message.Save.assert_not_called()
        self.client_class.assert_not_called()
        self.assertIn("Total calendar events in timeframe: 0", self.output.getvalue())
        self.assertNotIn("Total emails", self.output.getvalue())

    def test_calendar_category_added_during_request_is_preserved(self):
        appointment = event()
        self.calendar.Items = OutlookItems([appointment])

        def classify(**_kwargs):
            appointment.Categories = "Personal"
            return SimpleNamespace(
                choices={"category": SimpleNamespace(choice="Important", confidence=0.9)}
            )

        self.client.system_one.side_effect = classify
        self.assertEqual(self.run_categorization(process_calendar=True), 0)
        self.assertEqual(appointment.Categories, "Personal")
        appointment.Save.assert_not_called()
        self.assertIn("Categorized during model request; unchanged", self.output.getvalue())

    def test_category_added_before_event_request_is_preserved(self):
        appointment = event()
        first_appointment = event("First event")
        self.calendar.Items = OutlookItems([first_appointment, appointment])

        def classify(**_kwargs):
            appointment.Categories = "Personal"
            return SimpleNamespace(
                choices={"category": SimpleNamespace(choice="Important", confidence=0.9)}
            )

        self.client.system_one.side_effect = classify
        self.assertEqual(self.run_categorization(process_calendar=True), 0)
        self.client.system_one.assert_called_once()
        first_appointment.Save.assert_called_once_with()
        appointment.Save.assert_not_called()
        self.assertEqual(appointment.Categories, "Personal")

    def test_invalid_calendar_response_is_logged_without_saving(self):
        appointment = event()
        self.calendar.Items = OutlookItems([appointment])
        self.client.system_one.return_value.choices["category"].choice = "Unknown category"
        with self.assertLogs("outlook_categorization", level="ERROR") as logs:
            self.assertEqual(self.run_categorization(process_calendar=True), 1)
        appointment.Save.assert_not_called()
        self.assertIn("Calendar event", logs.output[0])
        self.assertIn("Start", logs.output[0])
        self.assertIn("End", logs.output[0])
        self.assertIn("validating the model response", logs.output[0])

    def test_save_failure_is_logged_and_other_calendar_events_continue(self):
        failed = event("Save failure")
        successful = event("Successful")
        failed.Save.side_effect = categorizer.pywintypes.com_error(
            -2147467259, "Save failed", None, None
        )
        self.calendar.Items = OutlookItems([failed, successful])
        with self.assertLogs("outlook_categorization", level="ERROR") as logs:
            self.assertEqual(self.run_categorization(process_calendar=True), 1)
        successful.Save.assert_called_once_with()
        self.assertIn("saving the selected category", logs.output[0])

    def test_calendar_access_failure_has_calendar_specific_message(self):
        self.namespace.GetDefaultFolder.side_effect = categorizer.pywintypes.com_error(
            -2147467259, "Cannot open Calendar", None, None
        )
        with self.assertRaisesRegex(RuntimeError, "Could not read Calendar events"):
            self.run_categorization(process_calendar=True)
        self.client_class.assert_not_called()


if __name__ == "__main__":
    unittest.main()
