"""
Royal Mail International 2026 carrier handlers.

Two variants share a single Royal Mail OBA portal:
  - Royal Mail International 2026        → Flats only (Ireland)
  - Royal Mail International 2026 - Ireland (P) → Letters only (Ireland)

Both use the standard carrier sheet layout (B3=carrier, B4=PO, data from row 8).
There is no manifest template — the manifest is downloaded from the portal.
The portal has a single form where both Flats and Letters data are entered together.
"""

import os
from datetime import datetime
from typing import Dict, Tuple, Optional
from dataclasses import dataclass

from openpyxl import load_workbook

from .base import BaseCarrier, ShipmentRecord, PlacementResult


@dataclass
class RoyalMailData:
    """Data extracted from a Royal Mail International carrier sheet.

    Weight fields store total weight in KG from the carrier sheet.
    The portal requires average item weight in grams, computed via
    avg_weight_grams().
    """
    po_number: str
    flats_items: int = 0
    flats_weight: float = 0.0      # total weight in KG
    letters_items: int = 0
    letters_weight: float = 0.0    # total weight in KG
    carrier_variant: str = ""      # 'flats' or 'letters'

    def avg_weight_grams(self, format_type: str) -> int:
        """Average item weight in grams for a format (Letters or Flats).

        The OBA portal requires this as a whole number in grams.
        """
        if format_type == 'letters' and self.letters_items > 0:
            return round(self.letters_weight * 1000 / self.letters_items)
        elif format_type == 'flats' and self.flats_items > 0:
            return round(self.flats_weight * 1000 / self.flats_items)
        return 0


class RoyalMailBaseCarrier(BaseCarrier):
    """
    Base handler for Royal Mail International 2026 carriers.

    Like Deutsche Post, this carrier has no manifest template.
    Data is extracted from the carrier sheet and submitted to the
    Royal Mail OBA portal, which generates the manifest.
    """

    carrier_name = ""
    template_filename = ""  # No template — portal generates the manifest
    expected_format = ""    # Subclasses set to 'Flats' or 'Letters'
    variant_key = ""        # Subclasses set to 'flats' or 'letters'

    # Ireland country name variations
    IRELAND_NAMES = {
        'ireland', 'republic of ireland', 'eire', 'ie',
        'ireland, republic of', 'roi',
    }

    def __init__(self):
        super().__init__()
        self.country_mapping = {
            'Republic of Ireland': 'Ireland',
            'Eire': 'Ireland',
            'IE': 'Ireland',
            'Ireland, Republic of': 'Ireland',
            'ROI': 'Ireland',
        }

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
        raise NotImplementedError("Royal Mail does not use place_record — use process_carrier_sheet")

    def process_carrier_sheet(
        self,
        carrier_sheet_path: str,
        output_dir: str,
        log_callback=None
    ) -> Tuple[str, RoyalMailData]:
        """
        Extract data from a Royal Mail carrier sheet and save to output dir.

        Args:
            carrier_sheet_path: Path to the original carrier sheet
            output_dir: Directory to save the processed sheet
            log_callback: Optional logging function

        Returns:
            (output_path, extracted_data)
        """
        def log(msg):
            if log_callback:
                log_callback(msg)

        wb = load_workbook(carrier_sheet_path, data_only=True)
        ws = wb.active

        # Extract metadata
        carrier_name = str(ws['B3'].value or "").strip()
        po_raw = ws['B4'].value
        po_number = str(int(po_raw)) if isinstance(po_raw, float) else str(po_raw or "")

        # Read data rows (header at row 8, data from row 9)
        headers = {}
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=8, column=col).value
            if val:
                headers[str(val).strip()] = col

        total_items = 0
        total_weight = 0.0
        non_ireland_countries = []
        unexpected_formats = []

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

            # Validate country is Ireland
            mapped_country = self.map_country(country)
            if mapped_country.lower() not in self.IRELAND_NAMES:
                non_ireland_countries.append(country)

            # Validate format matches expected
            normalised_format = self.normalise_format(format_val)
            if normalised_format != self.expected_format:
                unexpected_formats.append(format_val)

            # Sum totals
            try:
                items = int(items_val) if items_val not in (None, '', ' ') else 0
            except (ValueError, TypeError):
                items = 0

            try:
                weight = float(weight_val) if weight_val not in (None, '', ' ') else 0.0
            except (ValueError, TypeError):
                weight = 0.0

            total_items += items
            total_weight += weight
            row += 1

        # Log warnings
        if non_ireland_countries:
            unique = set(non_ireland_countries)
            log(f"  ⚠ Non-Ireland countries found: {', '.join(unique)}")

        if unexpected_formats:
            unique = set(unexpected_formats)
            log(f"  ⚠ Unexpected formats for {self.expected_format}-only carrier: {', '.join(unique)}")

        # Build result data
        data = RoyalMailData(
            po_number=po_number,
            carrier_variant=self.variant_key,
        )

        if self.variant_key == 'flats':
            data.flats_items = total_items
            data.flats_weight = round(total_weight, 3)
        else:
            data.letters_items = total_items
            data.letters_weight = round(total_weight, 3)

        # Save carrier sheet to output directory
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = carrier_name.replace(" ", "_").replace("/", "-")
        output_filename = f"{safe_name}_{po_number}_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        wb.save(output_path)
        wb.close()

        return output_path, data


class RoyalMailFlatsCarrier(RoyalMailBaseCarrier):
    """Handler for Royal Mail International 2026 (Flats only, Ireland)."""

    carrier_name = "Royal Mail International 2026"
    expected_format = "Flats"
    variant_key = "flats"


class RoyalMailLettersCarrier(RoyalMailBaseCarrier):
    """Handler for Royal Mail International 2026 - Ireland (P) (Letters only, Ireland)."""

    carrier_name = "Royal Mail International 2026 - Ireland (P)"
    expected_format = "Letters"
    variant_key = "letters"
