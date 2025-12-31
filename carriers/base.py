"""
Base carrier class defining the interface all carrier modules must implement.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional
import pandas as pd


@dataclass
class ShipmentRecord:
    """Standardised shipment record from carrier sheet."""
    country: str
    service: str  # 'Priority' or 'Economy'
    format: str   # 'Letters', 'Flats', 'Packets'
    items: int
    weight: float


@dataclass
class PlacementResult:
    """Result of placing a record in the manifest."""
    success: bool
    sheet_name: str = ""
    row: int = 0
    items_col: int = 0
    weight_col: int = 0
    error_message: str = ""


class BaseCarrier(ABC):
    """Abstract base class for carrier manifest handlers."""
    
    # Carrier identification
    carrier_name: str = ""
    template_filename: str = ""
    
    # Service type mapping from carrier sheet to internal representation
    service_map: Dict[str, str] = {}
    
    def __init__(self):
        self.country_mapping: Dict[str, str] = {}
        self.errors: List[str] = []
    
    @abstractmethod
    def build_country_index(self, workbook) -> Dict[str, dict]:
        """
        Build index mapping country names to their locations in the manifest.
        
        Returns dict like:
        {
            'France': {
                'Priority': {'sheet': 'Main Europe', 'row': 15, 'section': 'left'},
                'Economy': {'sheet': 'Main Europe', 'row': 15, 'section': 'left'}
            }
        }
        """
        pass
    
    @abstractmethod
    def get_cell_positions(self, country_info: dict, format_type: str) -> Tuple[int, int]:
        """
        Get the (items_column, weight_column) for a given format within a country row.
        
        Args:
            country_info: Dict with 'sheet', 'row', 'section' keys
            format_type: 'Letters', 'Flats', or 'Packets'
        
        Returns:
            Tuple of (items_column, weight_column)
        """
        pass
    
    @abstractmethod
    def set_metadata(self, workbook, po_number: str, shipment_date: str) -> None:
        """
        Set PO number and shipment date in the manifest.
        
        Args:
            workbook: openpyxl Workbook
            po_number: PO reference from carrier sheet
            shipment_date: Date string to set
        """
        pass
    
    def normalise_service(self, service: str) -> str:
        """Convert carrier sheet service name to 'Priority' or 'Economy'."""
        service_lower = service.lower().strip()
        if 'priority' in service_lower:
            return 'Priority'
        elif 'velociti' in service_lower:
            return 'Priority'
        elif 'economy' in service_lower:
            return 'Economy'
        return service
    
    def normalise_format(self, format_type: str) -> str:
        """Normalise format name to standard: Letters, Flats, Packets."""
        fmt_lower = format_type.lower().strip()
        if 'letter' in fmt_lower:
            return 'Letters'
        elif 'flat' in fmt_lower or 'boxable' in fmt_lower:
            return 'Flats'
        elif 'packet' in fmt_lower or 'non-boxable' in fmt_lower or 'nonboxable' in fmt_lower:
            return 'Packets'
        return format_type
    
    def map_country(self, carrier_country: str) -> Optional[str]:
        """Map carrier sheet country name to manifest country name."""
        if carrier_country in self.country_mapping:
            return self.country_mapping[carrier_country]
        return carrier_country
    
    def place_record(self, workbook, record: ShipmentRecord, country_index: dict) -> PlacementResult:
        """
        Place a shipment record into the manifest.
        
        Returns PlacementResult indicating success/failure.
        """
        manifest_country = self.map_country(record.country)
        
        if manifest_country not in country_index:
            return PlacementResult(
                success=False,
                error_message=f"Country not found in manifest: {record.country} (mapped to: {manifest_country})"
            )
        
        service = self.normalise_service(record.service)
        format_type = self.normalise_format(record.format)
        
        country_data = country_index[manifest_country]
        
        if service not in country_data:
            return PlacementResult(
                success=False,
                error_message=f"Service '{service}' not available for {manifest_country}"
            )
        
        location = country_data[service]
        sheet = workbook[location['sheet']]
        row = location['row']
        
        try:
            items_col, weight_col = self.get_cell_positions(location, format_type)
        except ValueError as e:
            return PlacementResult(
                success=False,
                error_message=str(e)
            )
        
        # Get current values (may be None, empty string, or already have data)
        current_items_raw = sheet.cell(row=row, column=items_col).value
        current_weight_raw = sheet.cell(row=row, column=weight_col).value
        
        # Convert to numeric, treating None/empty/non-numeric as 0
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
            sheet_name=location['sheet'],
            row=row,
            items_col=items_col,
            weight_col=weight_col
        )
