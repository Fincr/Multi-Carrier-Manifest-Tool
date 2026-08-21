"""
Royal Mail International 2026 carrier handler.

A single carrier sheet may contain Flats and Letters rows for any of the
supported destinations (see carriers/royalmail_countries.py). Data is
extracted and bucketed per destination and format, then submitted to the
Royal Mail OBA portal which generates the manifest. There is no manifest
template, and nothing is written to the output folder — the only artefact
is the order confirmation PDF the portal automation saves there.
"""

from typing import Dict, List, Tuple
from dataclasses import dataclass, field

from openpyxl import load_workbook

from .base import BaseCarrier, ShipmentRecord, PlacementResult
from .royalmail_countries import (
    CountryLine,
    resolve_country,
    supported_country_list,
)


# Formats covered by the OBA international products (PS5 / PS7)
SUPPORTED_FORMATS = ('Letters', 'Flats')


@dataclass
class RoyalMailData:
    """Data extracted from a Royal Mail International carrier sheet.

    One CountryLine per destination and format, each holding the sheet's
    item count and total weight in KG. The aggregate properties exist for
    the processing log; the portal works from the lines.
    """
    po_number: str
    lines: List[CountryLine] = field(default_factory=list)

    def _total(self, format_type: str, attr: str):
        return sum(
            getattr(line, attr) for line in self.lines
            if line.format_type == format_type
        )

    @property
    def letters_items(self) -> int:
        return self._total('Letters', 'items')

    @property
    def letters_weight(self) -> float:
        return round(self._total('Letters', 'weight_kg'), 3)

    @property
    def flats_items(self) -> int:
        return self._total('Flats', 'items')

    @property
    def flats_weight(self) -> float:
        return round(self._total('Flats', 'weight_kg'), 3)


class RoyalMailCarrier(BaseCarrier):
    """
    Handler for Royal Mail International 2026.

    Like Deutsche Post, this carrier has no manifest template.
    Data is extracted from the carrier sheet and submitted to the
    Royal Mail OBA portal, which generates the manifest.
    A single sheet may contain rows for several destinations and formats.
    """

    carrier_name = "Royal Mail International 2026"
    template_filename = ""  # No template — portal generates the manifest

    # Destination spellings live in royalmail_countries.resolve_country, which
    # both this parser and the portal automation share.

    def build_country_index(self, workbook) -> Dict[str, dict]:
        """Not used — Royal Mail has no template."""
        return {}

    def get_cell_positions(self, country_info: dict, format_type: str) -> Tuple[int, int]:
        """Not used — Royal Mail has no template."""
        raise NotImplementedError("Royal Mail does not use cell-based manifests")

    def set_metadata(self, workbook, po_number: str, shipment_date: str) -> None:
        """Not used — Royal Mail has no template."""
        pass

    def place_record(self, workbook, record: ShipmentRecord, country_index: dict) -> PlacementResult:
        """Not used — Royal Mail processes data via extract_data instead."""
        raise NotImplementedError("Royal Mail does not use place_record — use extract_data")

    def extract_data(self, carrier_sheet_path: str, log_callback=None) -> RoyalMailData:
        """
        Read the per-destination volumes a Royal Mail sheet declares.

        Nothing is written: OBA produces the manifest, so there is no
        output_dir argument and no copy of the sheet left behind.

        Args:
            carrier_sheet_path: Path to the original carrier sheet
            log_callback: Optional logging function

        Returns:
            The extracted RoyalMailData
        """
        def log(msg):
            if log_callback:
                log_callback(msg)

        wb = load_workbook(carrier_sheet_path, data_only=True)
        ws = wb.active

        # Extract metadata
        po_raw = ws['B4'].value
        po_number = str(int(po_raw)) if isinstance(po_raw, float) else str(po_raw or "")

        # Read data rows (header at row 8, data from row 9)
        headers = {}
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=8, column=col).value
            if val:
                headers[str(val).strip()] = col

        # Bucket by (destination, format) — one OBA order line each
        buckets: Dict[Tuple[str, str], CountryLine] = {}
        unsupported_countries = []
        unknown_formats = []

        row = 9
        while row <= ws.max_row:
            country_val = ws.cell(row=row, column=headers.get('Country', 1)).value
            if country_val is None or str(country_val).strip() == '':
                row += 1
                continue

            country = str(country_val).strip()
            format_val = str(ws.cell(row=row, column=headers.get('Format', 4)).value or "").strip()
            items_val = ws.cell(row=row, column=headers.get('Items', 5)).value
            weight_val = ws.cell(row=row, column=headers.get('Weight (KG)', 6)).value

            # Parse numeric values
            try:
                items = int(items_val) if items_val not in (None, '', ' ') else 0
            except (ValueError, TypeError):
                items = 0

            try:
                weight = float(weight_val) if weight_val not in (None, '', ' ') else 0.0
            except (ValueError, TypeError):
                weight = 0.0

            # A destination or format we cannot file is collected rather than
            # skipped: silently dropping a row, or folding it into another
            # country's line, means posting an order with the wrong volumes.
            resolved_country = resolve_country(country)
            normalised_format = self.normalise_format(format_val)

            if resolved_country is None:
                unsupported_countries.append(f"{country} (row {row})")
            elif normalised_format not in SUPPORTED_FORMATS:
                unknown_formats.append(f"{format_val or '(blank)'} (row {row})")
            else:
                key = (resolved_country, normalised_format)
                line = buckets.get(key)
                if line is None:
                    line = CountryLine(country=resolved_country, format_type=normalised_format)
                    buckets[key] = line
                line.items += items
                line.weight_kg = round(line.weight_kg + weight, 3)

            row += 1

        if unsupported_countries:
            raise ValueError(
                "Royal Mail carrier sheet contains destinations that cannot be "
                f"filed through OBA: {', '.join(unsupported_countries)}. "
                f"Supported: {supported_country_list()}."
            )

        if unknown_formats:
            raise ValueError(
                "Royal Mail carrier sheet contains formats that cannot be filed "
                f"through OBA: {', '.join(unknown_formats)}. "
                f"Supported: {', '.join(SUPPORTED_FORMATS)}."
            )

        # Build result data
        data = RoyalMailData(po_number=po_number, lines=list(buckets.values()))

        for line in data.lines:
            log(f"  {line.country} {line.format_type}: {line.items} items, "
                f"{line.weight_kg} kg ({line.avg_weight_grams}g avg)")

        wb.close()

        return data


