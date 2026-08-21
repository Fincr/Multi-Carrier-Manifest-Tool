"""Tests for discovering the carrier sheets in a folder chosen by the operator."""

import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.sheet_discovery import find_printable_sheets


class FindPrintableSheetsTests(unittest.TestCase):
    """
    Which files in a folder are candidates for the printer.

    Batch printing takes every workbook it finds, without reading B3: the
    operator points at a folder of paperwork and expects all of it to print.
    Batch *processing* layers carrier validation on top of the same scan, so
    both paths agree on what counts as a candidate file.
    """

    def setUp(self):
        self.folder = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.folder, ignore_errors=True)

    def _touch(self, name):
        path = os.path.join(self.folder, name)
        with open(path, 'w') as handle:
            handle.write('')
        return path

    def test_finds_both_excel_extensions(self):
        self._touch('Landmark.xlsx')
        self._touch('Legacy.xls')

        found = find_printable_sheets(self.folder)

        self.assertEqual(
            [os.path.basename(p) for p in found],
            ['Landmark.xlsx', 'Legacy.xls'],
        )

    def test_returns_absolute_paths(self):
        expected = self._touch('PostNord.xlsx')

        self.assertEqual(find_printable_sheets(self.folder), [expected])

    def test_skips_excel_lock_files(self):
        # Excel leaves a ~$ owner file behind while a workbook is open; it is
        # not paperwork and Excel cannot print it.
        self._touch('Asendia.xlsx')
        self._touch('~$Asendia.xlsx')

        found = find_printable_sheets(self.folder)

        self.assertEqual([os.path.basename(p) for p in found], ['Asendia.xlsx'])

    def test_ignores_files_that_are_not_workbooks(self):
        self._touch('Spring.xlsx')
        self._touch('notes.txt')
        self._touch('manifest.pdf')
        self._touch('data.csv')

        found = find_printable_sheets(self.folder)

        self.assertEqual([os.path.basename(p) for p in found], ['Spring.xlsx'])

    def test_matches_extensions_case_insensitively(self):
        # Sheets arriving by email are routinely upper-cased.
        self._touch('SHOUTING.XLSX')

        found = find_printable_sheets(self.folder)

        self.assertEqual([os.path.basename(p) for p in found], ['SHOUTING.XLSX'])

    def test_orders_results_so_the_print_queue_is_predictable(self):
        for name in ('charlie.xlsx', 'alpha.xlsx', 'bravo.xlsx'):
            self._touch(name)

        found = find_printable_sheets(self.folder)

        self.assertEqual(
            [os.path.basename(p) for p in found],
            ['alpha.xlsx', 'bravo.xlsx', 'charlie.xlsx'],
        )

    def test_ignores_subdirectories(self):
        # Only the folder the operator pointed at; a nested archive of last
        # month's sheets must not reach the printer.
        os.mkdir(os.path.join(self.folder, 'archive'))
        self._touch(os.path.join('archive', 'old.xlsx'))
        self._touch('current.xlsx')

        found = find_printable_sheets(self.folder)

        self.assertEqual([os.path.basename(p) for p in found], ['current.xlsx'])

    def test_a_directory_named_like_a_workbook_is_not_a_sheet(self):
        os.mkdir(os.path.join(self.folder, 'looks_like.xlsx'))

        self.assertEqual(find_printable_sheets(self.folder), [])

    def test_empty_folder_yields_nothing(self):
        self.assertEqual(find_printable_sheets(self.folder), [])

    def test_missing_folder_yields_nothing_rather_than_raising(self):
        # The GUI reports "no sheets found" for both cases, so discovery does
        # not need to distinguish them.
        missing = os.path.join(self.folder, 'does_not_exist')

        self.assertEqual(find_printable_sheets(missing), [])


if __name__ == '__main__':
    unittest.main()
