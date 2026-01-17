"""
United Business NZP ETOE carrier manifest handler.

Handles "Untracked Priority Mail" service for T&D Priority Manifest.
Single service type (Priority) with format-based columns (P/G/E format).
Simple single-row-per-country structure.
"""

from typing import Dict, Tuple
from .base import BaseCarrier, ShipmentRecord, PlacementResult


class UnitedBusinessNZPCarrier(BaseCarrier):
    """Handler for United Business NZP ETOE manifests (T&D Priority)."""
    
    carrier_name = "United Business NZP ETOE"
    template_filename = "UBL_CP_Pre_Alert_T_D-ETOE.xlsx"
    sheet_name = "Untracked Priority"
    
    # Column structure based on manifest template:
    # Row 4: P Format (G-H), G Format (I-J), E Format (K-L)
    # Row 5: Item | Kilo | Item | Kilo | Item | Kilo
    # P = Letters, G = Flats, E = Packets
    FORMAT_COLUMNS = {
        'Letters': (7, 8),    # G, H (P format - Items, Weight)
        'Flats': (9, 10),     # I, J (G format)
        'Packets': (11, 12),  # K, L (E format)
    }
    
    # Metadata cells
    DATE_CELL = 'B2'      # Next to "DATE:" in A2
    PO_CELL = 'B3'        # Next to "Citipost Job Reference" in A3
    
    # Data range
    DATA_START_ROW = 6
    DATA_END_ROW = 50
    
    def __init__(self):
        super().__init__()
        self.country_mapping = {
            # IST name -> Manifest name
            # Manifest uses 'Czechia' not 'Czech Republic'
            'Czech Republic': 'Czechia',
            
            # Taiwan variations
            'Taiwan': 'Taiwan, China',
            'Taiwan, Province of China': 'Taiwan, China',
            
            # Korea
            'Korea': 'South Korea',
            'Republic of Korea': 'South Korea',
            'Korea, Republic of': 'South Korea',
        }
        
        # Country row locations (built dynamically)
        self._country_locations: Dict[str, dict] = {}
    
    def build_country_index(self, workbook) -> Dict[str, dict]:
        """
        Build index mapping country names to their row locations.
        
        NZP ETOE has a simple structure: one row per country, single service (Priority).
        Countries are in column B (Destination), rows 6-50.
        """
        if self._country_locations:
            return self._country_locations
        
        sheet = workbook[self.sheet_name]
        
        for row in range(self.DATA_START_ROW, self.DATA_END_ROW + 1):
            country = sheet.cell(row=row, column=2).value  # Column B = Destination
            
            if not country:
                continue
            
            country_str = str(country).strip()
            
            # All entries are Priority service (Untracked Priority Mail)
            self._country_locations[country_str] = {
                'Priority': {
                    'sheet': self.sheet_name,
                    'row': row,
                }
            }
        
        return self._country_locations
    
    def get_cell_positions(self, country_info: dict, format_type: str) -> Tuple[int, int]:
        """Get columns for items and weight based on format."""
        if format_type not in self.FORMAT_COLUMNS:
            raise ValueError(f"Unknown format: {format_type}. Expected: Letters, Flats, or Packets")
        return self.FORMAT_COLUMNS[format_type]
    
    def set_metadata(self, workbook, po_number: str, shipment_date: str) -> None:
        """Set PO and date in the manifest header."""
        sheet = workbook[self.sheet_name]
        sheet[self.PO_CELL] = po_number
        sheet[self.DATE_CELL] = shipment_date
    
    def normalise_service(self, service: str) -> str:
        """NZP ETOE only has Priority service (Untracked Priority Mail)."""
        return 'Priority'


def get_carrier():
    """Factory function to get carrier instance."""
    return UnitedBusinessNZPCarrier()
