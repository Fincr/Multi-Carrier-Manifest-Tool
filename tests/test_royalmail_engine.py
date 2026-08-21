"""
Tests for how the engine handles a Royal Mail sheet.

Royal Mail's manifest is produced by the OBA portal, not by this tool, so
processing a Royal Mail sheet must leave the output folder untouched.
"""

import os
import sys
import shutil
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.engine import ManifestEngine
from tests.test_royalmail_parsing import build_sheet


class RoyalMailEngineTests(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.output_dir = os.path.join(self.tmp, 'output')
        os.makedirs(self.output_dir)
        self.sheet = os.path.join(self.tmp, 'carrier_sheet.xlsx')

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_engine(self, rows, **kwargs):
        build_sheet(self.sheet, rows, **kwargs)
        engine = ManifestEngine(self.tmp, self.output_dir)
        engine.set_log_callback(lambda msg: None)
        return engine.process_sheet(self.sheet)

    def test_nothing_is_written_to_the_output_folder(self):
        self.run_engine([('Ireland', 'Letters', 100, 5.0)])
        self.assertEqual(os.listdir(self.output_dir), [])

    def test_the_result_succeeds_with_no_output_file(self):
        results = self.run_engine([('Ireland', 'Letters', 100, 5.0)])
        self.assertEqual(len(results), 1)
        self.assertTrue(results[0].success)
        self.assertEqual(results[0].output_file, "")

    def test_the_extracted_destinations_reach_the_result(self):
        results = self.run_engine([
            ('Ireland', 'Letters', 100, 5.0),
            ('Canada', 'Flats', 20, 8.0),
        ])
        data = results[0].royalmail_data
        self.assertIsNotNone(data)
        self.assertEqual(
            sorted((l.country, l.format_type, l.items) for l in data.lines),
            [('Canada', 'Flats', 20), ('Ireland', 'Letters', 100)],
        )
        self.assertEqual(results[0].po_number, '123456')
        self.assertEqual(results[0].records_processed, 120)

    def test_an_unfileable_destination_fails_and_writes_nothing(self):
        results = self.run_engine([('Spain', 'Letters', 60, 3.0)])
        self.assertFalse(results[0].success)
        self.assertIn('Spain', ' '.join(results[0].errors))
        self.assertEqual(results[0].output_file, "")
        self.assertEqual(os.listdir(self.output_dir), [])


if __name__ == '__main__':
    unittest.main()
