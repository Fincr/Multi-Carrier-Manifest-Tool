"""
Metafora 2026 carrier manifest handlers (NZP and SPL variants).

List-style manifest: data rows are written dynamically starting at row 17,
one row per (country, format) combination. Both variants share the same
template but differ in the product code written to cell C6.

Format codes: P = Letters, G = Flats, E = Packets
"""

from datetime import datetime, timedelta
from typing import Dict, Tuple
from .base import BaseCarrier, ShipmentRecord, PlacementResult


class MetaforaBaseCarrier(BaseCarrier):
    """Base handler for Metafora 2026 manifests (shared logic for NZP and SPL)."""

    carrier_name = ""
    template_filename = "Metafora_Pre_Alert_2026.xlsx"
    sheet_name = "Pre-Alert"

    # Product code written to C6 - overridden by subclasses
    product_code = ""

    # Format code mapping: internal name -> manifest single-letter code
    FORMAT_CODES = {
        'Letters': 'P',
        'Flats': 'G',
        'Packets': 'E',
    }

    # Cell locations
    PO_CELL = 'C5'
    PRODUCT_CELL = 'C6'
    DATE_CELL = 'C8'

    # Data area
    DATA_START_ROW = 17
    COUNTRY_COL = 2   # B
    FORMAT_COL = 3     # C
    ITEMS_COL = 4      # D
    WEIGHT_COL = 5     # E
    AVG_COL = 6        # F (formula)

    def __init__(self):
        super().__init__()
        self.country_mapping = {
            # IST name -> Manifest name
            'Czech Republic': 'Czechia',

            # Taiwan variations
            'Taiwan': 'Taiwan, China',
            'Taiwan, Province of China': 'Taiwan, China',

            # Korea
            'Korea': 'South Korea',
            'Republic of Korea': 'South Korea',
            'Korea, Republic of': 'South Korea',

            # North Macedonia variations
            'Republic of North Macedonia': 'North Macedonia',
            'Macedonia': 'North Macedonia',

            # Vietnam
            'Vietnam': 'Viet Nam',

            # Ivory Coast
            'Cote d\'Ivoire': 'Ivory Coast',
            'Côte d\'Ivoire': 'Ivory Coast',
        }

        # Aggregated data: (country, format_code) -> [items, weight]
        self._aggregated_data: Dict[Tuple[str, str], list] = {}

    def _get_next_business_day(self, from_date: datetime) -> datetime:
        """Get the next business day (Monday-Friday) from the given date."""
        next_day = from_date + timedelta(days=1)
        while next_day.weekday() >= 5:  # Saturday=5, Sunday=6
            next_day += timedelta(days=1)
        return next_day

    def build_country_index(self, workbook) -> Dict[str, dict]:
        """Returns empty dict - Metafora has no pre-filled country rows."""
        return {}

    def get_cell_positions(self, country_info: dict, format_type: str) -> Tuple[int, int]:
        """Not used - Metafora writes rows dynamically."""
        raise NotImplementedError("Metafora uses dynamic row writing, not cell placement")

    def set_metadata(self, workbook, po_number: str, shipment_date: str) -> None:
        """Set PO number, product code, and expected arrival date."""
        sheet = workbook[self.sheet_name]
        sheet[self.PO_CELL] = po_number
        sheet[self.PRODUCT_CELL] = self.product_code

        # Expected arrival date = processed date + 1 business day
        today = datetime.now()
        arrival_date = self._get_next_business_day(today)
        sheet[self.DATE_CELL] = arrival_date.strftime("%Y-%m-%d")

    def normalise_service(self, service: str) -> str:
        """Metafora NZP/SPL are priority-only services."""
        return 'Priority'

    def place_record(self, workbook, record: ShipmentRecord, country_index: dict) -> PlacementResult:
        """
        Aggregate record into internal data store.

        Records are accumulated by (country, format_code) and written
        to the workbook later via flush_to_workbook().
        """
        format_type = self.normalise_format(record.format)
        format_code = self.FORMAT_CODES.get(format_type)

        if not format_code:
            return PlacementResult(
                success=False,
                error_message=f"Unknown format: {record.format} (normalised: {format_type})"
            )

        country = self.map_country(record.country)

        key = (country, format_code)
        if key in self._aggregated_data:
            self._aggregated_data[key][0] += record.items
            self._aggregated_data[key][1] += record.weight
        else:
            self._aggregated_data[key] = [record.items, record.weight]

        return PlacementResult(
            success=True,
            sheet_name=self.sheet_name,
            row=0,
            items_col=self.ITEMS_COL,
            weight_col=self.WEIGHT_COL
        )

    def flush_to_workbook(self, workbook) -> None:
        """Write all aggregated data rows into the manifest starting at row 17."""
        sheet = workbook[self.sheet_name]
        row = self.DATA_START_ROW

        for (country, format_code), (items, weight) in sorted(self._aggregated_data.items()):
            sheet.cell(row=row, column=self.COUNTRY_COL).value = country
            sheet.cell(row=row, column=self.FORMAT_COL).value = format_code
            sheet.cell(row=row, column=self.ITEMS_COL).value = items
            sheet.cell(row=row, column=self.WEIGHT_COL).value = round(weight, 3)
            sheet.cell(row=row, column=self.AVG_COL).value = f'=IFERROR(E{row}/D{row},0)'
            row += 1


class MetaforaNZPCarrier(MetaforaBaseCarrier):
    """Handler for Metafora 2026 - NZP manifests."""

    carrier_name = "Metafora 2026 - NZP"
    product_code = "Untracked_Priority_NZP"


class MetaforaSPLCarrier(MetaforaBaseCarrier):
    """Handler for Metafora 2026 - SPL manifests."""

    carrier_name = "Metafora 2026 - SPL"
    product_code = "Untracked_Priority_SPL"


def get_carrier_nzp():
    """Factory function to get NZP carrier instance."""
    return MetaforaNZPCarrier()


def get_carrier_spl():
    """Factory function to get SPL carrier instance."""
    return MetaforaSPLCarrier()
