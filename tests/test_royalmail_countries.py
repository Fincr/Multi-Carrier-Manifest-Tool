"""Tests for the Royal Mail OBA destination-country registry."""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from carriers.royalmail_countries import (
    SUPPORTED_COUNTRIES,
    CountryLine,
    resolve_country,
)


class ResolveCountryTests(unittest.TestCase):
    """Carrier sheets spell destinations inconsistently; all six must resolve."""

    def test_canonical_names_resolve_to_themselves(self):
        for name in ('Ireland', 'Netherlands', 'Italy', 'France', 'Germany', 'Canada'):
            self.assertEqual(resolve_country(name), name)

    def test_ireland_aliases_resolve(self):
        for alias in ('Republic of Ireland', 'Eire', 'IE', 'ROI', 'Ireland, Republic of'):
            self.assertEqual(resolve_country(alias), 'Ireland')

    def test_iso_codes_resolve(self):
        self.assertEqual(resolve_country('NL'), 'Netherlands')
        self.assertEqual(resolve_country('IT'), 'Italy')
        self.assertEqual(resolve_country('FR'), 'France')
        self.assertEqual(resolve_country('DE'), 'Germany')
        self.assertEqual(resolve_country('CA'), 'Canada')

    def test_resolution_ignores_case_and_surrounding_space(self):
        self.assertEqual(resolve_country('  netherlands  '), 'Netherlands')
        self.assertEqual(resolve_country('GERMANY'), 'Germany')

    def test_unsupported_country_resolves_to_none(self):
        self.assertIsNone(resolve_country('Spain'))
        self.assertIsNone(resolve_country(''))
        self.assertIsNone(resolve_country(None))


class CountrySpecTests(unittest.TestCase):
    """Each destination carries the OBA region it must be filed under."""

    def test_eu_destinations_use_european_union_region(self):
        for name in ('Ireland', 'Netherlands', 'Italy', 'France', 'Germany'):
            self.assertEqual(SUPPORTED_COUNTRIES[name].region, 'EUROPEAN UNION')

    def test_canada_uses_rest_of_world_region(self):
        self.assertEqual(SUPPORTED_COUNTRIES['Canada'].region, 'REST OF WORLD')

    def test_registry_order_fixes_the_order_of_lines_on_the_order_form(self):
        self.assertEqual(
            list(SUPPORTED_COUNTRIES),
            ['Ireland', 'Netherlands', 'Italy', 'France', 'Germany', 'Canada'],
        )


class CountryLineTests(unittest.TestCase):
    """OBA wants average item weight in whole grams, not the sheet's total KG."""

    def test_average_weight_converts_total_kg_to_grams_per_item(self):
        line = CountryLine('Ireland', 'Letters', items=1200, weight_kg=6.0)
        self.assertEqual(line.avg_weight_grams, 5)

    def test_average_weight_rounds_to_a_whole_gram(self):
        line = CountryLine('France', 'Flats', items=3, weight_kg=1.0)
        self.assertEqual(line.avg_weight_grams, 333)

    def test_average_weight_of_an_empty_line_is_zero_not_a_crash(self):
        line = CountryLine('Canada', 'Letters', items=0, weight_kg=0.0)
        self.assertEqual(line.avg_weight_grams, 0)


if __name__ == '__main__':
    unittest.main()
