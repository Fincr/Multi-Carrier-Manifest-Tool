"""
Tests for the Excel COM lifecycle behind printing.

Batch printing exposed a latent fault here: win32com's `Dispatch` returns an
already-running Excel from the Running Object Table, so once a print had called
`Quit()` the next print reattached to that quit-pending instance. Such an
instance still answers COM calls but silently refuses to open workbooks,
returning None instead of raising, which surfaced as an AttributeError about
NoneType several frames away.

Measured against 17 real manifests: `Dispatch` failed 10, `DispatchEx` failed
none.
"""

import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import gui

SHEET = 'C:' + chr(92) + 'sheets' + chr(92) + 'a.xlsx'
BROKEN = 'C:' + chr(92) + 'sheets' + chr(92) + 'broken.xlsx'


def fake_excel():
    """An Excel application mock whose workbook can be iterated for sheets."""
    excel = mock.MagicMock(name='Excel.Application')
    workbook = mock.MagicMock(name='Workbook')
    workbook.Sheets = [mock.MagicMock(name='Sheet1')]
    excel.Workbooks.Open.return_value = workbook
    return excel, workbook


class ExcelInstanceTests(unittest.TestCase):

    def test_a_private_excel_instance_is_created_for_each_print(self):
        """
        A print must not adopt an Excel instance it did not create.

        `Dispatch` would hand back a previous print's quit-pending instance —
        and, worse, the operator's own open Excel session, which this code then
        sets DisplayAlerts=False on and Quits.
        """
        excel, _ = fake_excel()

        with mock.patch('win32com.client.DispatchEx', return_value=excel) as dispatch_ex, \
             mock.patch('win32com.client.Dispatch') as plain_dispatch:
            success, message = gui.print_excel_workbook(SHEET, 'KT02')

        self.assertTrue(success, message)
        dispatch_ex.assert_called_once_with("Excel.Application")
        plain_dispatch.assert_not_called()

    def test_excel_is_quit_after_a_successful_print(self):
        excel, workbook = fake_excel()

        with mock.patch('win32com.client.DispatchEx', return_value=excel):
            gui.print_excel_workbook(SHEET, 'KT02')

        workbook.Close.assert_called_once()
        excel.Quit.assert_called_once()


class UnopenedWorkbookTests(unittest.TestCase):
    """
    A workbook Excel declines to open must be reported as such.

    Excel returns None here rather than raising. Left unchecked it travelled
    into the print helper and failed as "'NoneType' object has no attribute
    'ActiveSheet'", which tells the operator nothing about which file or why.
    """

    def test_a_workbook_excel_will_not_open_is_reported_plainly(self):
        excel, _ = fake_excel()
        excel.Workbooks.Open.return_value = None

        with mock.patch('win32com.client.DispatchEx', return_value=excel):
            success, message = gui.print_excel_workbook(BROKEN, 'KT02')

        self.assertFalse(success)
        self.assertNotIn('NoneType', message)
        self.assertIn('broken.xlsx', message)

    def test_excel_is_still_quit_when_the_workbook_never_opened(self):
        # Otherwise a failed print leaks an Excel process per attempt.
        excel, _ = fake_excel()
        excel.Workbooks.Open.return_value = None

        with mock.patch('win32com.client.DispatchEx', return_value=excel):
            gui.print_excel_workbook(BROKEN, 'KT02')

        excel.Quit.assert_called_once()


if __name__ == '__main__':
    unittest.main()
