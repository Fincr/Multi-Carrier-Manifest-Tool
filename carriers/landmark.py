"""
Landmark Global carrier handler.

Landmark uses a CSV upload format where:
- Line 1: CONTRACT_NR,PRODUCT_CODE,DEPOSIT_DATE,DEPOSIT_DAY_PART
- Line 2+: DEST_COUNTRY(ISO-2),FORMAT(P/G/E),WEIGHT(KG),PIECES

Economy (12SL03) and Priority (12SL02) must be in separate files.
"""

from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from datetime import datetime, timedelta
import os
import pandas as pd

from .base import BaseCarrier, ShipmentRecord, PlacementResult


@dataclass 
class LandmarkOrderLine:
    """A single order line for Landmark upload file."""
    iso_code: str
    format_code: str  # P, G, E
    weight_kg: float
    pieces: int
    product_code: str  # 12SL03 (Economy) or 12SL02 (Priority)


class LandmarkCarrier(BaseCarrier):
    """Handler for Landmark Global upload files."""
    
    carrier_name = "Landmark Global"
    template_filename = "UploadCodeList_-_Citipost.xls"  # Used for ISO codes, not as template
    
    # Static configuration
    CONTRACT_NR = "BPI/2024/00011075"
    DEPOSIT_DAY_PART = "PM"
    
    # Service to product code mapping
    SERVICE_TO_PRODUCT = {
        'Priority': '12SL02',
        'Economy': '12SL03',
    }
    
    # Format mapping (internal -> upload code)
    FORMAT_CODES = {
        'Letters': 'P',
        'Flats': 'G',
        'Packets': 'E',
    }
    
    def __init__(self):
        super().__init__()
        self._order_lines: List[LandmarkOrderLine] = []
        self._po_number = ""
        self._deposit_date = ""
        self._file_date = ""
        self._iso_map: Dict[str, str] = {}
        self._iso_codes_loaded = False
    
    def load_iso_codes(self, filepath: str) -> bool:
        """
        Load ISO-2 country codes from the Landmark code list.
        
        Args:
            filepath: Path to UploadCodeList_-_Citipost.xls
            
        Returns:
            True if loaded successfully
        """
        try:
            df_iso = pd.read_excel(filepath, sheet_name=1)
            
            for _, row in df_iso.iterrows():
                if pd.notna(row['NAME']) and pd.notna(row['ISO_CODE']):
                    name = str(row['NAME']).strip().lower()
                    code = str(row['ISO_CODE']).strip()
                    self._iso_map[name] = code
            
            # Add common variations
            self._add_country_variations()
            self._iso_codes_loaded = True
            return True
            
        except Exception as e:
            self.errors.append(f"Failed to load ISO codes: {e}")
            return False
    
    def _add_country_variations(self) -> None:
        """Add common country name variations to the ISO map."""
        variations = {
            'hong kong': 'HK',
            'hong-kong': 'HK',
            'turkey': 'TR',
            'türkiye': 'TR',
            'usa': 'US',
            'united states of america': 'US',
            'uk': 'GB',
            'united kingdom': 'GB',
            'great britain': 'GB',
            'south korea': 'KR',
            'north korea': 'KP',
            'czech republic': 'CZ',
            'czechia': 'CZ',
            'uae': 'AE',
            'russia': 'RU',
            'russian federation': 'RU',
            'vietnam': 'VN',
            'viet nam': 'VN',
        }
        for name, code in variations.items():
            if name not in self._iso_map:
                self._iso_map[name] = code
    
    def get_iso_code(self, country: str) -> Optional[str]:
        """Get ISO-2 code for a country name."""
        country_lower = country.lower().strip()
        return self._iso_map.get(country_lower)
    
    def build_country_index(self, workbook) -> Dict[str, dict]:
        """
        Landmark doesn't use a pre-built country index.
        We build order lines dynamically. Return empty dict.
        """
        return {}
    
    def get_cell_positions(self, country_info: dict, format_type: str) -> Tuple[int, int]:
        """Not used for Landmark - we build order lines instead."""
        raise NotImplementedError("Landmark uses order line generation, not cell placement")
    
    def _get_next_working_day(self, from_date: datetime) -> datetime:
        """
        Get the next working day (Monday-Friday) from the given date.
        
        Args:
            from_date: The starting date
            
        Returns:
            The next working day
        """
        next_day = from_date + timedelta(days=1)
        # Skip weekends: Saturday (5) and Sunday (6)
        while next_day.weekday() >= 5:
            next_day += timedelta(days=1)
        return next_day
    
    def set_metadata(self, workbook, po_number: str, shipment_date: str) -> None:
        """Store PO number and calculate deposit date (next working day)."""
        self._po_number = po_number
        
        # Always use today as base, then calculate next working day
        today = datetime.now()
        next_working_day = self._get_next_working_day(today)
        
        self._deposit_date = next_working_day.strftime("%d/%m/%Y")
        self._file_date = today.strftime("%Y%m%d")
    
    def place_record(self, workbook, record: ShipmentRecord, country_index: dict) -> PlacementResult:
        """
        Convert a shipment record to a Landmark order line.
        Instead of placing in cells, we accumulate order lines.
        """
        # Check ISO codes are loaded
        if not self._iso_codes_loaded:
            return PlacementResult(
                success=False,
                error_message="ISO codes not loaded. Call load_iso_codes() first."
            )
        
        # Get ISO code
        iso_code = self.get_iso_code(record.country)
        if not iso_code:
            return PlacementResult(
                success=False,
                error_message=f"No ISO code mapping for: {record.country}"
            )
        
        # Get service/product code
        service = self.normalise_service(record.service)
        product_code = self.SERVICE_TO_PRODUCT.get(service)
        if not product_code:
            return PlacementResult(
                success=False,
                error_message=f"Unknown service type: {record.service}"
            )
        
        # Get format code
        format_type = self.normalise_format(record.format)
        format_code = self.FORMAT_CODES.get(format_type)
        if not format_code:
            return PlacementResult(
                success=False,
                error_message=f"Unknown format type: {record.format}"
            )
        
        # Create order line
        order_line = LandmarkOrderLine(
            iso_code=iso_code,
            format_code=format_code,
            weight_kg=round(record.weight, 3),
            pieces=record.items,
            product_code=product_code,
        )
        
        self._order_lines.append(order_line)
        
        return PlacementResult(
            success=True,
            sheet_name="Upload",
            row=len(self._order_lines) + 1,
            items_col=4,
            weight_col=3,
        )
    
    def write_upload_files(self, output_dir: str) -> List[str]:
        """
        Write accumulated order lines to Landmark CSV upload files.
        
        Generates separate files for Economy (12SL03) and Priority (12SL02).
        
        Args:
            output_dir: Directory to save the CSV files
            
        Returns:
            List of generated file paths
        """
        files_created = []
        
        # Group by product code
        economy_lines = [line for line in self._order_lines if line.product_code == '12SL03']
        priority_lines = [line for line in self._order_lines if line.product_code == '12SL02']

        # Write Economy file
        if economy_lines:
            filename = f"Landmark_Economy_{self._file_date}_{self._po_number}.csv"
            filepath = os.path.join(output_dir, filename)
            self._write_csv(filepath, '12SL03', economy_lines)
            files_created.append(filepath)
        
        # Write Priority file
        if priority_lines:
            filename = f"Landmark_Priority_{self._file_date}_{self._po_number}.csv"
            filepath = os.path.join(output_dir, filename)
            self._write_csv(filepath, '12SL02', priority_lines)
            files_created.append(filepath)
        
        return files_created
    
    def _write_csv(self, filepath: str, product_code: str, lines: List[LandmarkOrderLine]) -> None:
        """Write a single CSV upload file."""
        with open(filepath, 'w', newline='') as f:
            # Header row (pipe-separated)
            f.write(f"{self.CONTRACT_NR}|{product_code}|{self._deposit_date}|{self.DEPOSIT_DAY_PART}||{self._po_number}|\n")
            # Data rows (pipe-separated)
            for line in lines:
                f.write(f"{line.iso_code}|{line.format_code}|{line.weight_kg}|{line.pieces}\n")
    
    def clear_order_lines(self) -> None:
        """Clear accumulated order lines for fresh processing."""
        self._order_lines = []
        self._po_number = ""
        self._deposit_date = ""
        self._file_date = ""
    
    def get_order_lines(self) -> List[LandmarkOrderLine]:
        """Return accumulated order lines."""
        return self._order_lines
    
    def get_summary(self) -> Dict[str, dict]:
        """Get a summary of accumulated order lines by service type."""
        economy_lines = [line for line in self._order_lines if line.product_code == '12SL03']
        priority_lines = [line for line in self._order_lines if line.product_code == '12SL02']
        
        summary = {}
        
        if economy_lines:
            summary['Economy'] = {
                'product_code': '12SL03',
                'rows': len(economy_lines),
                'total_pieces': sum(line.pieces for line in economy_lines),
                'total_weight_kg': round(sum(line.weight_kg for line in economy_lines), 3),
            }

        if priority_lines:
            summary['Priority'] = {
                'product_code': '12SL02',
                'rows': len(priority_lines),
                'total_pieces': sum(line.pieces for line in priority_lines),
                'total_weight_kg': round(sum(line.weight_kg for line in priority_lines), 3),
            }
        
        return summary
