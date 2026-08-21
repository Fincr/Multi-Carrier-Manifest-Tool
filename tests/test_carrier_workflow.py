"""Tests for the shared carrier-workflow decisions used by both processing paths."""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.carrier_workflow import prints_output_file


class PrintsOutputFileTests(unittest.TestCase):
    """
    Whether the local output file is the thing to send to the printer.

    Single-sheet and batch mode must agree, so both ask this function
    rather than each keeping its own list of exceptions.
    """

    def test_a_populated_template_is_printed_here(self):
        self.assertTrue(prints_output_file('Asendia 2026'))
        self.assertTrue(prints_output_file('PostNord'))

    def test_deutsche_post_prints_its_carrier_sheet_here(self):
        # The sheet with (EMB) Manifest removed IS Deutsche Post's paperwork.
        self.assertTrue(prints_output_file('Deutsche Post'))

    def test_spring_defers_printing_to_its_portal(self):
        self.assertFalse(prints_output_file('Spring'))

    def test_landmark_defers_printing_to_its_portal(self):
        self.assertFalse(prints_output_file('Landmark Global'))

    def test_royal_mail_defers_printing_to_its_portal(self):
        # Royal Mail writes no local file at all — OBA generates the manifest.
        self.assertFalse(prints_output_file('Royal Mail International 2026'))

    def test_the_carrier_name_is_matched_loosely(self):
        self.assertFalse(prints_output_file('  royal mail international 2026  '))
        self.assertFalse(prints_output_file('SPRING GDS'))

    def test_an_unknown_or_missing_carrier_still_prints(self):
        # Defaulting to printing keeps a new carrier's paperwork from vanishing.
        self.assertTrue(prints_output_file('Some New Carrier'))
        self.assertTrue(prints_output_file(''))
        self.assertTrue(prints_output_file(None))


if __name__ == '__main__':
    unittest.main()
