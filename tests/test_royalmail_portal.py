"""Tests for the browser-free logic inside the Royal Mail OBA portal automation."""

import asyncio
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from carriers.royalmail_countries import CountryLine
from carriers.royalmail_portal import (
    OBA_LOGIN_URL,
    PRODUCT_CODES,
    RoyalMailPortalInput,
    _create_order,
    _find_royalmail_page,
    match_option_value,
)


class MatchOptionValueTests(unittest.TestCase):
    """
    OBA's dropdown labels are not something we can verify from here, so we
    match against whatever the portal actually rendered instead of hardcoding
    a guess like 'NETHERLANDS (THE)'.
    """

    def test_exact_label_matches(self):
        options = [('1', 'FRANCE'), ('2', 'GERMANY')]
        self.assertEqual(match_option_value(options, 'Germany', 'country'), '2')

    def test_label_with_a_suffix_matches_on_its_leading_words(self):
        options = [('7', 'IRELAND (REPUBLIC OF)'), ('8', 'ITALY')]
        self.assertEqual(match_option_value(options, 'Ireland', 'country'), '7')

    def test_leading_match_is_preferred_over_a_label_that_merely_contains_it(self):
        options = [('1', 'NORTHERN IRELAND'), ('2', 'IRELAND (REPUBLIC OF)')]
        self.assertEqual(match_option_value(options, 'Ireland', 'country'), '2')

    def test_exact_label_is_preferred_over_a_longer_leading_match(self):
        options = [('1', 'IRELAND (REPUBLIC OF)'), ('2', 'IRELAND')]
        self.assertEqual(match_option_value(options, 'Ireland', 'country'), '2')

    def test_matching_ignores_case_and_extra_whitespace(self):
        options = [('5', '  netherlands   (the) ')]
        self.assertEqual(match_option_value(options, 'Netherlands', 'country'), '5')

    def test_region_labels_match_the_same_way(self):
        options = [('A', 'EUROPEAN UNION'), ('B', 'REST OF WORLD')]
        self.assertEqual(match_option_value(options, 'REST OF WORLD', 'region'), 'B')

    def test_no_match_raises_and_reports_what_the_portal_offered(self):
        options = [('1', 'FRANCE'), ('2', 'GERMANY')]
        with self.assertRaises(LookupError) as ctx:
            match_option_value(options, 'Canada', 'country')
        message = str(ctx.exception)
        self.assertIn('Canada', message)
        self.assertIn('FRANCE', message)
        self.assertIn('GERMANY', message)
        self.assertIn('country', message)

    def test_an_ambiguous_match_raises_rather_than_guessing(self):
        options = [('1', 'CONGO (BRAZZAVILLE)'), ('2', 'CONGO (KINSHASA)')]
        with self.assertRaises(LookupError) as ctx:
            match_option_value(options, 'Congo', 'country')
        self.assertIn('ambiguous', str(ctx.exception).lower())

    def test_label_containing_the_name_matches_when_nothing_leads_with_it(self):
        options = [('3', 'THE NETHERLANDS'), ('4', 'BELGIUM')]
        self.assertEqual(match_option_value(options, 'Netherlands', 'country'), '3')

    def test_blank_placeholder_options_are_ignored(self):
        options = [('', ''), ('', '-- please select --'), ('9', 'CANADA')]
        self.assertEqual(match_option_value(options, 'Canada', 'country'), '9')


class ProductCodeTests(unittest.TestCase):

    def test_each_format_maps_to_its_oba_product_code(self):
        self.assertEqual(PRODUCT_CODES['Letters'], 'PS5')
        self.assertEqual(PRODUCT_CODES['Flats'], 'PS7')


class OrderedLinesTests(unittest.TestCase):
    """
    Line order on the OBA form is fixed so a confirmation PDF always reads
    the same way, whatever order the carrier sheets arrived in.
    """

    def test_lines_follow_registry_country_order(self):
        portal_input = RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Canada', 'Letters', 10, 1.0),
            CountryLine('Ireland', 'Letters', 20, 2.0),
            CountryLine('Germany', 'Letters', 30, 3.0),
        ])
        self.assertEqual(
            [line.country for line in portal_input.ordered_lines],
            ['Ireland', 'Germany', 'Canada'],
        )

    def test_letters_precede_flats_within_a_country(self):
        portal_input = RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Italy', 'Flats', 5, 1.0),
            CountryLine('Italy', 'Letters', 5, 1.0),
        ])
        self.assertEqual(
            [line.format_type for line in portal_input.ordered_lines],
            ['Letters', 'Flats'],
        )

    def test_lines_with_no_items_are_dropped(self):
        portal_input = RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Ireland', 'Letters', 0, 0.0),
            CountryLine('France', 'Flats', 7, 1.0),
        ])
        self.assertEqual(
            [(line.country, line.format_type) for line in portal_input.ordered_lines],
            [('France', 'Flats')],
        )

    def test_has_any_is_false_when_every_line_is_empty(self):
        empty = RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Ireland', 'Letters', 0, 0.0),
        ])
        self.assertFalse(empty.has_any)
        populated = RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Ireland', 'Letters', 1, 0.1),
        ])
        self.assertTrue(populated.has_any)


class MergeTests(unittest.TestCase):
    """
    Batch mode processes several carrier sheets into one OBA order, so their
    volumes have to combine without double-counting a destination.
    """

    def test_matching_country_and_format_are_summed(self):
        target = RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Ireland', 'Letters', 100, 5.0),
        ])
        target.merge(RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Ireland', 'Letters', 50, 2.5),
        ]))
        self.assertEqual(len(target.lines), 1)
        self.assertEqual(target.lines[0].items, 150)
        self.assertAlmostEqual(target.lines[0].weight_kg, 7.5)

    def test_a_new_destination_is_added_as_its_own_line(self):
        target = RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Ireland', 'Letters', 100, 5.0),
        ])
        target.merge(RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Italy', 'Flats', 20, 4.0),
        ]))
        self.assertEqual(
            sorted((l.country, l.format_type) for l in target.lines),
            [('Ireland', 'Letters'), ('Italy', 'Flats')],
        )

    def test_same_country_different_format_stays_a_separate_line(self):
        target = RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Ireland', 'Letters', 100, 5.0),
        ])
        target.merge(RoyalMailPortalInput(po_number='PO1', lines=[
            CountryLine('Ireland', 'Flats', 10, 3.0),
        ]))
        self.assertEqual(len(target.lines), 2)

    def test_merging_fills_in_a_missing_po_number(self):
        target = RoyalMailPortalInput(po_number='', lines=[])
        target.merge(RoyalMailPortalInput(po_number='PO9', lines=[
            CountryLine('Ireland', 'Letters', 1, 0.1),
        ]))
        self.assertEqual(target.po_number, 'PO9')

    def test_merging_does_not_overwrite_an_existing_po_number(self):
        target = RoyalMailPortalInput(po_number='PO1', lines=[])
        target.merge(RoyalMailPortalInput(po_number='PO9', lines=[]))
        self.assertEqual(target.po_number, 'PO1')


class UnfileableLineTests(unittest.TestCase):
    """
    A line the portal cannot file is rejected before the browser is touched,
    with a message that names the problem. These calls pass None for the page
    and frame precisely to prove nothing is touched before the check.
    """

    def create_order(self, portal_input):
        return asyncio.run(
            _create_order(None, None, portal_input, lambda msg: None, 30000)
        )

    def test_an_unknown_destination_is_reported_by_name(self):
        success, message = self.create_order(RoyalMailPortalInput(
            po_number='PO1', lines=[CountryLine('Spain', 'Letters', 10, 1.0)],
        ))
        self.assertFalse(success)
        self.assertIn('Spain', message)
        self.assertIn('Ireland', message)

    def test_an_unknown_format_is_reported_by_name(self):
        success, message = self.create_order(RoyalMailPortalInput(
            po_number='PO1', lines=[CountryLine('Ireland', 'Packets', 10, 1.0)],
        ))
        self.assertFalse(success)
        self.assertIn('Packets', message)
        self.assertIn('Letters', message)

    def test_an_order_with_nothing_to_declare_is_refused(self):
        success, message = self.create_order(RoyalMailPortalInput(
            po_number='PO1', lines=[CountryLine('Ireland', 'Letters', 0, 0.0)],
        ))
        self.assertFalse(success)
        self.assertIn('No Royal Mail volumes', message)



class FakePage:
    def __init__(self, url):
        self.url = url

    async def goto(self, url, **kwargs):
        self.url = url


class FakeContext:
    """
    A browser context whose tab list can change between polls: each read of
    `pages` returns the next snapshot, then keeps returning the last one.
    """

    def __init__(self, *snapshots, can_open=True):
        self.snapshots = [list(s) for s in snapshots] or [[]]
        self.can_open = can_open
        self.opened = []

    @property
    def pages(self):
        if len(self.snapshots) > 1:
            return self.snapshots.pop(0)
        return self.snapshots[0] + self.opened

    async def new_page(self):
        if not self.can_open:
            raise RuntimeError('cannot open a tab')
        page = FakePage('about:blank')
        self.opened.append(page)
        return page


class FakeBrowser:
    def __init__(self, *contexts):
        self.contexts = list(contexts)


class FindRoyalMailPageTests(unittest.TestCase):
    """
    Every run now starts a fresh Edge, and the tab Edge was told to open is
    not guaranteed to be there the moment the debug port answers. The lookup
    waits for it, and opens the login page itself rather than giving up.
    """

    def find(self, browser):
        return asyncio.run(
            _find_royalmail_page(browser, lambda msg: None, wait_seconds=1, poll_seconds=0)
        )

    def test_an_open_royal_mail_tab_is_used(self):
        tab = FakePage('https://www.royalmail.com/login')
        page, error = self.find(FakeBrowser(FakeContext([FakePage('edge://newtab/'), tab])))
        self.assertIs(page, tab)
        self.assertEqual(error, '')

    def test_a_tab_that_navigates_there_after_connecting_is_used(self):
        tab = FakePage('https://www.royalmail.com/login')
        context = FakeContext([FakePage('about:blank')], [FakePage('about:blank')], [tab])
        page, error = self.find(FakeBrowser(context))
        self.assertIs(page, tab)

    def test_the_login_page_is_opened_when_no_tab_ever_arrives(self):
        context = FakeContext([FakePage('edge://sync-confirmation-dialog/')])
        page, error = self.find(FakeBrowser(context))
        self.assertIsNotNone(page)
        self.assertEqual(page.url, OBA_LOGIN_URL)
        self.assertEqual(error, '')

    def test_failure_to_open_a_tab_reports_the_tabs_that_were_seen(self):
        context = FakeContext([FakePage('edge://sync-confirmation-dialog/')], can_open=False)
        page, error = self.find(FakeBrowser(context))
        self.assertIsNone(page)
        self.assertIn('edge://sync-confirmation-dialog/', error)

    def test_a_browser_with_no_contexts_is_reported_rather_than_crashing(self):
        page, error = self.find(FakeBrowser())
        self.assertIsNone(page)
        self.assertIn('Royal Mail', error)


if __name__ == '__main__':
    unittest.main()
