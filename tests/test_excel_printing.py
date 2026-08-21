"""Tests for what happens to an open workbook on its way to the printer."""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.excel_printing import print_workbook, should_abandon_batch_print


class FakePageSetup:
    """
    Stands in for Excel's PageSetup, recording the order of assignments.

    Order matters: while `Zoom` holds a number Excel ignores FitToPages
    entirely, so `Zoom = False` has to land first or the fit silently does
    nothing.
    """

    def __init__(self, raises=False):
        self._raises = raises
        self.assignments = []

    def __setattr__(self, name, value):
        if name.startswith('_') or name == 'assignments':
            object.__setattr__(self, name, value)
            return
        if self._raises:
            raise Exception("PageSetup unavailable on this sheet type")
        self.assignments.append((name, value))
        object.__setattr__(self, name, value)


class FakeSheet:
    def __init__(self, name, page_setup_raises=False):
        self.Name = name
        self.PageSetup = FakePageSetup(raises=page_setup_raises)
        self.print_calls = []

    def PrintOut(self, **kwargs):
        self.print_calls.append(kwargs)


class FakeWorkbook:
    def __init__(self, sheets, active=0):
        self.Sheets = sheets
        self.ActiveSheet = sheets[active]
        self.print_calls = []

    def PrintOut(self, **kwargs):
        self.print_calls.append(kwargs)


class PrintTargetTests(unittest.TestCase):
    """Which object gets sent to the printer."""

    def test_whole_workbook_prints_by_default(self):
        # The manifest print path relies on every tab coming out.
        sheets = [FakeSheet('Summary'), FakeSheet('Main Europe')]
        wb = FakeWorkbook(sheets)

        print_workbook(wb)

        self.assertEqual(len(wb.print_calls), 1)
        self.assertEqual([s.print_calls for s in sheets], [[], []])

    def test_active_sheet_only_prints_that_sheet_alone(self):
        # A carrier sheet's reference tabs are not paperwork.
        sheets = [FakeSheet('Carrier Sheet'), FakeSheet('Lists')]
        wb = FakeWorkbook(sheets, active=0)

        print_workbook(wb, active_sheet_only=True)

        self.assertEqual(len(sheets[0].print_calls), 1)
        self.assertEqual(sheets[1].print_calls, [])
        self.assertEqual(wb.print_calls, [])

    def test_active_sheet_is_honoured_when_it_is_not_the_first_tab(self):
        sheets = [FakeSheet('Notes'), FakeSheet('Carrier Sheet')]
        wb = FakeWorkbook(sheets, active=1)

        print_workbook(wb, active_sheet_only=True)

        self.assertEqual(sheets[0].print_calls, [])
        self.assertEqual(len(sheets[1].print_calls), 1)


class PrinterSelectionTests(unittest.TestCase):
    """Where the pages come out."""

    def test_named_printer_is_passed_through(self):
        wb = FakeWorkbook([FakeSheet('manifest')])

        print_workbook(wb, printer_name=r'\print01.citipost.co.uk\KT02')

        self.assertEqual(
            wb.print_calls,
            [{'ActivePrinter': r'\print01.citipost.co.uk\KT02'}],
        )

    def test_no_printer_name_uses_the_windows_default(self):
        # Passing ActivePrinter=None would be an error, not a default.
        wb = FakeWorkbook([FakeSheet('manifest')])

        print_workbook(wb, printer_name=None)

        self.assertEqual(wb.print_calls, [{}])

    def test_named_printer_reaches_a_single_sheet_print(self):
        sheets = [FakeSheet('Carrier Sheet')]
        wb = FakeWorkbook(sheets)

        print_workbook(wb, printer_name='KT02', active_sheet_only=True)

        self.assertEqual(sheets[0].print_calls, [{'ActivePrinter': 'KT02'}])


class FitToWidthTests(unittest.TestCase):
    """Columns must not run off the side of the page."""

    def test_every_sheet_is_fitted_when_printing_the_workbook(self):
        sheets = [FakeSheet('Summary'), FakeSheet('ROW')]
        wb = FakeWorkbook(sheets)

        print_workbook(wb)

        for sheet in sheets:
            self.assertIs(sheet.PageSetup.Zoom, False)
            self.assertEqual(sheet.PageSetup.FitToPagesWide, 1)

    def test_rows_are_left_unconstrained_so_long_sheets_stay_readable(self):
        sheet = FakeSheet('manifest')
        wb = FakeWorkbook([sheet])

        print_workbook(wb)

        self.assertEqual(sheet.PageSetup.FitToPagesTall, 32767)

    def test_zoom_is_disabled_before_the_fit_is_requested(self):
        sheet = FakeSheet('manifest')
        wb = FakeWorkbook([sheet])

        print_workbook(wb)

        names = [name for name, _ in sheet.PageSetup.assignments]
        self.assertLess(names.index('Zoom'), names.index('FitToPagesWide'))

    def test_only_the_active_sheet_is_fitted_when_printing_it_alone(self):
        sheets = [FakeSheet('Carrier Sheet'), FakeSheet('Lists')]
        wb = FakeWorkbook(sheets, active=0)

        print_workbook(wb, active_sheet_only=True)

        self.assertEqual(sheets[0].PageSetup.FitToPagesWide, 1)
        self.assertEqual(sheets[1].PageSetup.assignments, [])

    def test_a_sheet_that_rejects_page_setup_does_not_stop_the_print(self):
        # Chart sheets raise on PageSetup; the data tabs must still come out.
        sheets = [FakeSheet('Chart', page_setup_raises=True), FakeSheet('Data')]
        wb = FakeWorkbook(sheets)

        print_workbook(wb)

        self.assertEqual(sheets[1].PageSetup.FitToPagesWide, 1)
        self.assertEqual(len(wb.print_calls), 1)


class AbandonBatchPrintTests(unittest.TestCase):
    """
    When a failed sheet means giving up on the rest of the folder.

    A jammed or offline printer fails every remaining sheet, and pushing forty
    doomed workbooks through Excel wastes minutes at the machine. But a single
    corrupt file is no reason to abandon the sheets behind it, because the
    operator would have to re-run the folder and reprint what already came out.

    The dividing line: if nothing has printed yet, the printer itself is
    suspect. Once one sheet has come out, the printer works and later failures
    belong to the files.
    """

    def test_stops_when_the_very_first_sheet_fails(self):
        # Nothing has printed, so the printer is the likely culprit.
        self.assertTrue(should_abandon_batch_print(failure_count=1, attempted=1))

    def test_stops_when_nothing_has_printed_after_several_attempts(self):
        self.assertTrue(should_abandon_batch_print(failure_count=4, attempted=4))

    def test_continues_once_a_sheet_has_printed(self):
        # The printer demonstrably works; this sheet is the problem.
        self.assertFalse(should_abandon_batch_print(failure_count=1, attempted=2))

    def test_continues_even_when_most_sheets_have_failed(self):
        # One success proves the queue is alive, so the rest are worth trying.
        self.assertFalse(should_abandon_batch_print(failure_count=9, attempted=10))

    def test_an_early_success_protects_the_whole_rest_of_the_batch(self):
        # Sheet 1 printed, then 2 and 3 failed: keep going.
        self.assertFalse(should_abandon_batch_print(failure_count=2, attempted=3))


if __name__ == '__main__':
    unittest.main()
