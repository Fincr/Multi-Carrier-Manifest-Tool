"""
Tests for the batch print loop's accounting.

Drives the real `run_batch_print` with only the Excel COM call replaced, so
the loop, the early-stop policy and the counts handed to the summary are the
production ones.
"""

import os
import sys
import unittest
from types import SimpleNamespace
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import gui

# The UNC queue the tool prints to, spelled as config holds it.
PRINTER = '\\\\print01.citipost.co.uk\\KT02'

SHEET = 'C:\\sheets\\a.xlsx'


class StubApp:
    """
    Enough of the GUI for the batch print loop to run headlessly.

    `after` runs the callback inline rather than queueing it: the real one
    marshals onto the Tk event loop, and calling straight through keeps the
    assertions about ordering honest without a Tk instance.
    """

    def __init__(self):
        self.config = SimpleNamespace(printer_name=PRINTER)
        self.root = self
        self.status_var = SimpleNamespace(set=lambda text: None)
        self.logged = []
        self.completion = None

    def after(self, _delay, func, *args):
        func(*args)

    def log(self, message):
        self.logged.append(message)

    def on_batch_print_complete(self, total, attempted, failures):
        self.completion = SimpleNamespace(total=total, attempted=attempted, failures=failures)


def drive(sheets, outcomes):
    """Run the batch print loop, returning results per sheet from `outcomes`."""
    app = StubApp()
    with mock.patch.object(gui, 'print_excel_workbook', side_effect=outcomes) as printer:
        gui.ManifestToolApp.run_batch_print(app, sheets)
    return app, printer


class BatchPrintAccountingTests(unittest.TestCase):

    def test_every_sheet_printing_is_reported_as_all_sent(self):
        sheets = [f'C:\\sheets\\{n}.xlsx' for n in ('a', 'b', 'c')]

        app, _ = drive(sheets, [(True, 'Sent to printer: KT02')] * 3)

        self.assertEqual(app.completion.total, 3)
        self.assertEqual(app.completion.attempted, 3)
        self.assertEqual(app.completion.failures, [])

    def test_each_sheet_is_printed_as_its_active_sheet_to_the_configured_printer(self):
        app, printer = drive([SHEET], [(True, 'ok')])

        printer.assert_called_once_with(SHEET, PRINTER, active_sheet_only=True)

    def test_a_single_bad_sheet_does_not_stop_the_ones_behind_it(self):
        sheets = ['a.xlsx', 'b.xlsx', 'c.xlsx']
        outcomes = [(True, 'ok'), (False, 'corrupt file'), (True, 'ok')]

        app, printer = drive(sheets, outcomes)

        self.assertEqual(printer.call_count, 3)
        self.assertEqual(app.completion.attempted, 3)
        self.assertEqual(app.completion.failures, [('b.xlsx', 'corrupt file')])

    def test_a_dead_printer_stops_the_batch_after_the_first_sheet(self):
        sheets = [f'{n}.xlsx' for n in ('a', 'b', 'c', 'd')]
        outcomes = [(False, 'Print failed: printer offline')] * 4

        app, printer = drive(sheets, outcomes)

        self.assertEqual(printer.call_count, 1)
        self.assertEqual(app.completion.total, 4)
        self.assertEqual(app.completion.attempted, 1)
        self.assertEqual(len(app.completion.failures), 1)

    def test_unattempted_sheets_are_not_counted_as_printed(self):
        # The bug this guards: reporting printed as total - failures claimed
        # three sheets had come out when the batch stopped after the first.
        sheets = [f'{n}.xlsx' for n in ('a', 'b', 'c', 'd')]

        app, _ = drive(sheets, [(False, 'printer offline')] * 4)

        printed = app.completion.attempted - len(app.completion.failures)
        unattempted = app.completion.total - app.completion.attempted
        self.assertEqual(printed, 0)
        self.assertEqual(unattempted, 3)

    def test_the_operator_is_told_why_the_batch_stopped(self):
        sheets = ['a.xlsx', 'b.xlsx']

        app, _ = drive(sheets, [(False, 'printer offline')] * 2)

        self.assertTrue(
            any('not attempted' in message for message in app.logged),
            f"expected an explanation in the log, got: {app.logged}",
        )

    def test_a_failure_on_the_last_sheet_needs_no_early_stop(self):
        # Nothing remains, so the policy is not consulted and the batch simply
        # finishes with one failure recorded.
        app, printer = drive(['a.xlsx'], [(False, 'corrupt file')])

        self.assertEqual(printer.call_count, 1)
        self.assertEqual(app.completion.attempted, 1)
        self.assertEqual(app.completion.total, 1)


if __name__ == '__main__':
    unittest.main()
