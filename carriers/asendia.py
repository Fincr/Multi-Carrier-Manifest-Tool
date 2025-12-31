"""
Asendia UK Business Mail manifest handler.
"""

from typing import Dict, Tuple
from .base import BaseCarrier


class AsendiaCarrier(BaseCarrier):
    """Handler for Asendia UK Business Mail manifests."""
    
    carrier_name = "Asendia 2026"
    template_filename = "Asendia_UK_Business_2026_Mail_Manifest.xlsx"
    
    # Row where ROW (Rest of World) section starts in manifest
    ROW_SECTION_START = 76
    
    # EU countries use format-specific columns
    EU_FORMAT_COLUMNS = {
        'Letters': (5, 6),    # E, F
        'Flats': (10, 11),    # J, K
        'Packets': (15, 16),  # O, P
    }
    
    # ROW countries use left/right sections
    ROW_COLUMNS = {
        'left': (5, 6),       # E, F
        'right': (15, 16),    # O, P
    }
    
    def __init__(self):
        super().__init__()
        self.country_mapping = {
            # IST name -> Manifest name
            'United States of America': 'United States',
            'Aland Islands': 'Åland Islands',
            'South Korea': 'Korea, Republic of',
            'Taiwan': 'Taiwan, Province of China',
            'Bolivia': 'Bolivia, Plurinational State of',
            'Reunion': 'Réunion',
            'Saint Martin': 'Saint Martin (French part)',
            'Sint Maarten': 'Sint Maarten (Dutch part)',
            'Falkland Islands': 'Falkland Islands (Malvinas)',
            'Saint Helena': 'Saint Helena, Ascension and Tristan da Cunha',
            'Svalbard': 'Svalbard and Jan Mayen',
            'Palestine': 'Palestine, State of',
            'Democratic Republic of the Congo': 'Democratic Republic of Congo',
            'Antigua and Barbuda': 'Antigua And Barbuda',
            'Saint Kitts and Nevis': 'Saint Kitts And Nevis',
            'Sao Tome and Principe': 'Sao Tome And Principe',
            'Trinidad and Tobago': 'Trinidad And Tobago',
            'Turks and Caicos Islands': 'Turks And Caicos Islands',
            'Czech Republic': 'Czechia',
            'Côte d\'Ivoire': 'Cote d\'Ivoire',
            'Cote d\'Ivoire': 'Cote d\'Ivoire',
            'Ivory Coast': 'Cote d\'Ivoire',
            'Vietnam': 'Viet Nam',
            'Laos': 'Lao People\'s Democratic Republic',
            'Russia': 'Russian Federation',
            'Venezuela': 'Venezuela, Bolivarian Republic of',
            'Iran': 'Iran, Islamic Republic of',
            'Syria': 'Syrian Arab Republic',
            'Tanzania': 'Tanzania, United Republic of',
            'Micronesia': 'Micronesia, Federated States of',
            'Moldova': 'Moldova, Republic of',
            'Brunei': 'Brunei Darussalam',
            'North Korea': 'Korea, Democratic People\'s Republic of',
            'Republic of the Congo': 'Congo',
            'Virgin Islands (US)': 'Virgin Islands, U.S.',
            'Virgin Islands (British)': 'Virgin Islands, British',
            'Bonaire': 'Bonaire, Sint Eustatius and Saba',
            'Eswatini': 'Swaziland',
            'Myanmar (Burma)': 'Myanmar',
            'East Timor': 'Timor-Leste',
            'Macedonia': 'North Macedonia',
            'Cape Verde': 'Cabo Verde',
        }
    
    def build_country_index(self, workbook) -> Dict[str, dict]:
        """Build index for both Priority and Non-Priority manifest sheets."""
        index = {}
        
        for sheet_name in ['Priority Manifest', 'Non-Priority Manifest']:
            service = 'Priority' if 'Priority' == sheet_name.split()[0] else 'Economy'
            sheet = workbook[sheet_name]
            
            skip_values = {
                'Asendia UK Business Mail', 'Valid from', 'Work Required',
                'Customer Name', 'Office Check', 'Subtotal', 'TOTAL',
                'Shipment date', 'Customer Ref', 'PO', None
            }
            
            current_section = 'EU'
            for row in range(13, sheet.max_row + 1):
                country = sheet.cell(row=row, column=2).value
                
                if country is None:
                    continue
                    
                country_str = str(country).strip()
                
                if country_str in skip_values or country_str.startswith('Valid from'):
                    continue
                
                if row >= self.ROW_SECTION_START:
                    current_section = 'ROW'
                
                # Determine if this is left or right column in ROW section
                section = 'left'
                if current_section == 'ROW':
                    # Check if there's content in column K (right section country name)
                    right_country = sheet.cell(row=row, column=11).value
                    if right_country and str(right_country).strip():
                        # Add right-side country too
                        right_name = str(right_country).strip()
                        if right_name not in skip_values:
                            if right_name not in index:
                                index[right_name] = {}
                            index[right_name][service] = {
                                'sheet': sheet_name,
                                'row': row,
                                'section': 'right',
                                'type': 'ROW'
                            }
                
                if country_str not in index:
                    index[country_str] = {}
                
                index[country_str][service] = {
                    'sheet': sheet_name,
                    'row': row,
                    'section': section,
                    'type': current_section
                }
        
        return index
    
    def get_cell_positions(self, country_info: dict, format_type: str) -> Tuple[int, int]:
        """Get columns for items and weight based on section type and format."""
        if country_info['type'] == 'EU':
            if format_type not in self.EU_FORMAT_COLUMNS:
                raise ValueError(f"Unknown format: {format_type}")
            return self.EU_FORMAT_COLUMNS[format_type]
        else:
            # ROW section - uses left/right columns, all formats consolidated
            return self.ROW_COLUMNS[country_info['section']]
    
    def set_metadata(self, workbook, po_number: str, shipment_date: str) -> None:
        """Set PO and date in both manifest sheets."""
        for sheet_name in ['Priority Manifest', 'Non-Priority Manifest']:
            sheet = workbook[sheet_name]
            sheet['I6'] = po_number
            sheet['N6'] = shipment_date


class Asendia2025Carrier(AsendiaCarrier):
    """Handler for Asendia UK Business Mail 2025 manifests.
    
    Identical structure to 2026, just different template file.
    """
    
    carrier_name = "Asendia 2025"
    template_filename = "Asendia_UK_Business_Mail_2025.xlsx"
