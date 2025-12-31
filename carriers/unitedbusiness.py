"""
United Business Limited (UBL) ADS Mail manifest handler.

Handles "Untracked Economy Mail" service for Eastern Europe and Asia regions.
Single service type with format-based columns (Letters, Flats, Packets).
Some countries have multiple weight bands (China, Russia, Ukraine).
"""

from typing import Dict, Tuple, Optional
from .base import BaseCarrier, ShipmentRecord, PlacementResult


class UnitedBusinessCarrier(BaseCarrier):
    """Handler for United Business Limited ADS Mail manifests."""
    
    carrier_name = "United Business ADS"
    template_filename = "United_Business.xlsx"
    
    # Column structure for UBL manifest
    # Row 6-7 headers: Country | Item Weight | Letters (Items/Weight) | Flats (Items/Weight) | Packets (Items/Weight) | Total (Items/Weight)
    # Data columns:     A      |     B       |   C   |   D   |   E   |   F   |   G   |   H   |   I   |   J
    FORMAT_COLUMNS = {
        'Letters': (3, 4),    # C, D (Items, Weight)
        'Flats': (5, 6),      # E, F
        'Packets': (7, 8),    # G, H
    }
    
    # Metadata cells (merged cells in manifest)
    PO_CELL = 'F1'      # Ref. Nr. value cell
    DATE_CELL = 'F2'    # Date value cell
    
    def __init__(self):
        super().__init__()
        self.country_mapping = {
            # IST name -> UBL Manifest name
            # Note: UBL manifest has some spelling variations
            'Bosnia and Herzegovina': 'Bosnia & Herzegovina',
            'Bosnia-Herzegovina': 'Bosnia & Herzegovina',
            'Czech Republic': 'Czech Republic',
            'Czechia': 'Czech Republic',
            'Moldova': 'Moldova Republic',
            'Republic of Moldova': 'Moldova Republic',
            'Moldova, Republic of': 'Moldova Republic',
            'North Macedonia': 'Macedonia',
            'Republic of North Macedonia': 'Macedonia',
            'Serbia and Montenegro': 'Serbia & Montenegro',
            'Serbia': 'Serbia & Montenegro',
            'Montenegro': 'Serbia & Montenegro',
            'Myanmar': 'Myanmar',
            'Myanmar (Burma)': 'Myanmar',
            'Taiwan': 'Taiwan',
            'Taiwan, Province of China': 'Taiwan',
            'Russian Federation': 'Russia',
            'Vietnam': 'Vietnam',
            'Viet Nam': 'Vietnam',
            'Kyrgyzstan': 'Kyrgystan',  # Note: manifest has typo
            'Afghanistan': 'Afganistan',  # Note: manifest has typo
            'Azerbaijan': 'Azerbajan',  # Note: manifest has typo
        }
        
        # Country locations will be built dynamically
        self._country_locations: Dict[str, dict] = {}
    
    def _parse_weight_range(self, weight_str: str) -> Tuple[int, int]:
        """
        Parse weight range string to (min_grams, max_grams).
        Examples: '0g-2000g', '51-200g', '0g-50g', '21g-2000g'
        """
        if not weight_str:
            return (0, 2000)  # Default full range
        
        # Clean string and parse
        weight_str = str(weight_str).strip().lower().replace(' ', '')
        
        # Handle different formats
        parts = weight_str.replace('g', '').split('-')
        if len(parts) == 2:
            try:
                min_g = int(parts[0])
                max_g = int(parts[1])
                return (min_g, max_g)
            except ValueError:
                pass
        
        # Default full range
        return (0, 2000)
    
    def _find_weight_row(self, country_data: dict, weight_kg: float) -> int:
        """
        Find the appropriate row for a given weight from country's weight ranges.
        
        Args:
            country_data: Dict with 'rows' list of {row, min_g, max_g}
            weight_kg: Weight in kilograms
            
        Returns:
            Row number for the matching weight range
        """
        weight_g = weight_kg * 1000  # Convert to grams
        
        for row_info in country_data['rows']:
            min_g, max_g = row_info['min_g'], row_info['max_g']
            if min_g <= weight_g <= max_g:
                return row_info['row']
        
        # If no exact match, use the last (largest) weight range
        return country_data['rows'][-1]['row']
    
    def build_country_index(self, workbook) -> Dict[str, dict]:
        """
        Build index mapping country names to their locations in the manifest.
        
        UBL has a single sheet with countries in column A and weight ranges in column B.
        Some countries (China, Russia, Ukraine) have multiple rows for different weight ranges.
        """
        if self._country_locations:
            return self._country_locations
        
        sheet = workbook.active
        
        # Scan from row 9 (first data row) to find all countries
        for row in range(9, sheet.max_row + 1):
            country = sheet.cell(row=row, column=1).value
            weight_range = sheet.cell(row=row, column=2).value
            
            if not country:
                continue
            
            country_str = str(country).strip()
            weight_range_str = str(weight_range).strip() if weight_range else '0g-2000g'
            
            # Parse weight range
            min_g, max_g = self._parse_weight_range(weight_range_str)
            
            # Initialize or append to country entry
            if country_str not in self._country_locations:
                self._country_locations[country_str] = {
                    'sheet': sheet.title,
                    'rows': []
                }
            
            self._country_locations[country_str]['rows'].append({
                'row': row,
                'weight_range': weight_range_str,
                'min_g': min_g,
                'max_g': max_g
            })
        
        return self._country_locations
    
    def get_cell_positions(self, country_info: dict, format_type: str) -> Tuple[int, int]:
        """Get columns for items and weight based on format."""
        if format_type not in self.FORMAT_COLUMNS:
            raise ValueError(f"Unknown format: {format_type}. Expected: Letters, Flats, or Packets")
        return self.FORMAT_COLUMNS[format_type]
    
    def set_metadata(self, workbook, po_number: str, shipment_date: str) -> None:
        """Set PO and date in the manifest header."""
        sheet = workbook.active
        sheet[self.PO_CELL] = po_number
        sheet[self.DATE_CELL] = shipment_date
    
    def place_record(self, workbook, record: ShipmentRecord, country_index: dict) -> PlacementResult:
        """
        Place a shipment record into the manifest.
        
        Handles UBL's weight-range based row selection for countries
        like China, Russia, Ukraine that have multiple weight bands.
        """
        manifest_country = self.map_country(record.country)
        
        if manifest_country not in country_index:
            return PlacementResult(
                success=False,
                error_message=f"Country not found in manifest: {record.country} (mapped to: {manifest_country})"
            )
        
        format_type = self.normalise_format(record.format)
        country_data = country_index[manifest_country]
        
        # For UBL, all records are "Economy" (single service type)
        # Find the appropriate row based on average weight per item
        if record.items > 0:
            avg_weight_kg = record.weight / record.items
        else:
            avg_weight_kg = record.weight if record.weight > 0 else 0.1  # Default
        
        row = self._find_weight_row(country_data, avg_weight_kg)
        sheet_name = country_data['sheet']
        
        try:
            items_col, weight_col = self.get_cell_positions({}, format_type)
        except ValueError as e:
            return PlacementResult(
                success=False,
                error_message=str(e)
            )
        
        sheet = workbook[sheet_name]
        
        # Get current values
        current_items_raw = sheet.cell(row=row, column=items_col).value
        current_weight_raw = sheet.cell(row=row, column=weight_col).value
        
        try:
            current_items = int(current_items_raw) if current_items_raw not in (None, '', ' ') else 0
        except (ValueError, TypeError):
            current_items = 0
        
        try:
            current_weight = float(current_weight_raw) if current_weight_raw not in (None, '', ' ') else 0.0
        except (ValueError, TypeError):
            current_weight = 0.0
        
        # Add to existing values
        sheet.cell(row=row, column=items_col).value = current_items + record.items
        sheet.cell(row=row, column=weight_col).value = round(current_weight + record.weight, 3)
        
        return PlacementResult(
            success=True,
            sheet_name=sheet_name,
            row=row,
            items_col=items_col,
            weight_col=weight_col
        )
    
    def normalise_service(self, service: str) -> str:
        """UBL only has one service: Economy (Untracked Economy Mail)."""
        return 'Economy'
