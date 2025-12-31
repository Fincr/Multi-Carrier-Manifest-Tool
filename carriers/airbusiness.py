"""
Air Business manifest handler.

Air Business is an Ireland-only carrier with a simple fixed structure:
- Single sheet "Ireland Mail"
- Three format rows at fixed positions
- Weight in column K, Quantity in column L

Cell mapping:
  Letters: Row 16 - K16 (weight), L16 (items)
  Flats:   Row 19 - K19 (weight), L19 (items)
  Packets: Row 22 - K22 (weight), L22 (items)
  
Metadata:
  D4: Customer/PO reference
  D6: Shipment date
"""

from typing import Dict, Tuple
from .base import BaseCarrier, ShipmentRecord, PlacementResult


class AirBusinessCarrier(BaseCarrier):
    """Handler for Air Business Ireland manifests."""
    
    carrier_name = "Air Business"
    template_filename = "Air_Business_Ireland.xlsx"
    
    # Fixed row positions for each format
    FORMAT_ROWS = {
        'Letters': 16,
        'Flats': 19,
        'Packets': 22,
    }
    
    # Column positions (same for all formats)
    WEIGHT_COL = 11   # K
    ITEMS_COL = 12    # L
    
    def __init__(self):
        super().__init__()
        # Air Business only handles Ireland - include common variations
        self.country_mapping = {
            'Republic of Ireland': 'Ireland',
            'Eire': 'Ireland',
            'IE': 'Ireland',
        }
    
    def build_country_index(self, workbook) -> Dict[str, dict]:
        """
        Build country index - for Air Business this is just Ireland
        with all services routing to the same location.
        """
        return {
            'Ireland': {
                'Priority': {'sheet': 'Ireland Mail', 'row': None, 'section': 'fixed'},
                'Economy': {'sheet': 'Ireland Mail', 'row': None, 'section': 'fixed'},
            }
        }
    
    def get_cell_positions(self, country_info: dict, format_type: str) -> Tuple[int, int]:
        """
        Get (items_column, weight_column) for a format.
        Air Business uses fixed rows per format, same columns for all.
        """
        if format_type not in self.FORMAT_ROWS:
            raise ValueError(f"Unknown format for Air Business: {format_type}. Expected: Letters, Flats, or Packets")
        return (self.ITEMS_COL, self.WEIGHT_COL)
    
    def set_metadata(self, workbook, po_number: str, shipment_date: str) -> None:
        """Air Business manifest has no PO/date fields - no-op."""
        pass
    
    def place_record(self, workbook, record: ShipmentRecord, country_index: dict) -> PlacementResult:
        """
        Place a shipment record into the Air Business manifest.
        
        Overrides base class to handle the fixed-row structure.
        """
        # Validate country is Ireland
        manifest_country = self.map_country(record.country)
        if manifest_country != 'Ireland':
            return PlacementResult(
                success=False,
                error_message=f"Air Business only handles Ireland, got: {record.country}"
            )
        
        # Normalise format
        format_type = self.normalise_format(record.format)
        
        if format_type not in self.FORMAT_ROWS:
            return PlacementResult(
                success=False,
                error_message=f"Unknown format: {record.format} (normalised: {format_type})"
            )
        
        # Get the fixed row for this format
        row = self.FORMAT_ROWS[format_type]
        sheet = workbook['Ireland Mail']
        
        # Get current values
        current_items_raw = sheet.cell(row=row, column=self.ITEMS_COL).value
        current_weight_raw = sheet.cell(row=row, column=self.WEIGHT_COL).value
        
        # Convert to numeric
        try:
            current_items = int(current_items_raw) if current_items_raw not in (None, '', ' ') else 0
        except (ValueError, TypeError):
            current_items = 0
        
        try:
            current_weight = float(current_weight_raw) if current_weight_raw not in (None, '', ' ') else 0.0
        except (ValueError, TypeError):
            current_weight = 0.0
        
        # Add values
        sheet.cell(row=row, column=self.ITEMS_COL).value = current_items + record.items
        sheet.cell(row=row, column=self.WEIGHT_COL).value = round(current_weight + record.weight, 3)
        
        return PlacementResult(
            success=True,
            sheet_name='Ireland Mail',
            row=row,
            items_col=self.ITEMS_COL,
            weight_col=self.WEIGHT_COL
        )
