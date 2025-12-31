"""
PostNord Business Mail manifest handler.
"""

from typing import Dict, Tuple
from .base import BaseCarrier


class PostNordCarrier(BaseCarrier):
    """Handler for PostNord Business Mail manifests."""
    
    carrier_name = "PostNord"
    template_filename = "PostNord.xlsx"
    
    # Main Europe & Rest of Europe column structure (same for both)
    # Priority: Letters C-D, Boxable E-F, Nonboxable G-H
    # Economy: Letters I-J, Boxable K-L, Nonboxable M-N
    EUROPE_COLUMNS = {
        'Priority': {
            'Letters': (3, 4),    # C, D
            'Flats': (5, 6),      # E, F (Boxable)
            'Packets': (7, 8),    # G, H (Nonboxable)
        },
        'Economy': {
            'Letters': (9, 10),   # I, J
            'Flats': (11, 12),    # K, L (Boxable)
            'Packets': (13, 14),  # M, N (Nonboxable)
        }
    }
    
    # ROW sheets: left section (A-G), right section (H-N)
    # Both use: Letters B-C/I-J, Flats D-E/K-L, Packets F-G/M-N
    ROW_COLUMNS = {
        'left': {
            'Letters': (2, 3),    # B, C
            'Flats': (4, 5),      # D, E
            'Packets': (6, 7),    # F, G
        },
        'right': {
            'Letters': (9, 10),   # I, J
            'Flats': (11, 12),    # K, L
            'Packets': (13, 14),  # M, N
        }
    }
    
    def __init__(self):
        super().__init__()
        self.country_mapping = {
            # Carrier sheet name -> PostNord manifest name
            'United States of America': 'USA',
            'United States': 'USA',
            'Czech Republic': 'Czech Rep',
            'Czechia': 'Czech Rep',
            'Bosnia and Herzegovina': 'Bosnia Her.',
            'Bosnia-Herzegovina': 'Bosnia Her.',
            'North Macedonia': 'Macedonia',
            'Republic of North Macedonia': 'Macedonia',
            'Ivory Coast': 'Ivory Coast',
            "Côte d'Ivoire": 'Ivory Coast',
            "Cote d'Ivoire": 'Ivory Coast',
            'Democratic Republic of the Congo': 'Congo, Dem. Rep.',
            'Democratic Republic of Congo': 'Congo, Dem. Rep.',
            'Republic of the Congo': 'Congo, Rep. of',
            'Congo': 'Congo, Rep. of',
            'Central African Republic': 'Central African Rep.',
            'Laos': 'Laos, Rep. of',
            "Lao People's Democratic Republic": 'Laos, Rep. of',
            'South Korea': 'Korea',
            'Korea, Republic of': 'Korea',
            'Vietnam': 'Vietnam',
            'Viet Nam': 'Vietnam',
            'Brunei': 'Brunei Darussalam',
            'UAE': 'UAE',
            'United Arab Emirates': 'UAE',
            'Antigua and Barbuda': 'Antigua & Barbuda',
            'Trinidad and Tobago': 'Trinidad & Tobago',
            'Saint Kitts and Nevis': 'St. Kitts & Nevis',
            'Saint Lucia': 'St. Lucia',
            'Saint Vincent': 'St. Vincent',
            'Saint Vincent and the Grenadines': 'St. Vincent',
            'Turks and Caicos Islands': 'Turks & Caicos',
            'Turks and Caicos': 'Turks & Caicos',
            'Réunion': 'Reunion',
            'Norway ': 'Norway',  # Handle trailing space in manifest
            'Romania ': 'Romania',
            'Slovakia ': 'Slovakia',
            ' Montenegro': 'Montenegro',  # Handle leading space
            'Russia ': 'Russia',
            'Serbia ': 'Serbia',
            'Turkey ': 'Turkey',
            'Ukraine ': 'Ukraine',
            'Kazakhstan ': 'Kazakhstan',
            'Russian Federation': 'Russia',
            'Eswatini': 'Swaziland',
            'Taiwan, Province of China': 'Taiwan',
            'Bolivia, Plurinational State of': 'Bolivia',
            'Venezuela, Bolivarian Republic of': 'Venezuela',
            'Iran, Islamic Republic of': 'Iran',  # Not in PostNord
            'Moldova, Republic of': 'Moldova',
            'Tanzania, United Republic of': 'Tanzania',
            'Micronesia, Federated States of': 'Micronesia',  # Not in PostNord
        }
        
        # Build the country index structure
        self._country_locations = self._build_static_index()
    
    def _build_static_index(self) -> Dict[str, dict]:
        """
        Build static country location index.
        PostNord has fixed positions - we hardcode them for reliability.
        """
        index = {}
        
        # Main Europe - rows 8-34, all use left section columns
        main_europe_countries = {
            'Austria': 8, 'Belgium': 9, 'Croatia': 10, 'Czech Rep': 11,
            'Denmark': 12, 'Estonia': 13, 'Finland': 14, 'France': 15,
            'Germany': 16, 'Greece': 17, 'Hungary': 18, 'Iceland': 19,
            'Ireland': 20, 'Italy': 21, 'Latvia': 22, 'Lithuania': 23,
            'Luxembourg': 24, 'Netherlands': 25, 'Norway': 26, 'Poland': 27,
            'Portugal': 28, 'Romania': 29, 'Slovakia': 30, 'Slovenia': 31,
            'Spain': 32, 'Sweden': 33, 'Switzerland': 34
        }
        for country, row in main_europe_countries.items():
            index[country] = {
                'Priority': {'sheet': 'Main Europe', 'row': row, 'section': 'europe'},
                'Economy': {'sheet': 'Main Europe', 'row': row, 'section': 'europe'}
            }
        
        # Rest of Europe - rows 8-27
        rest_europe_countries = {
            'Albania': 8, 'Armenia': 9, 'Azerbaijan': 10, 'Belarus': 11,
            'Bosnia Her.': 12, 'Bulgaria': 13, 'Cyprus': 14, 'Georgia': 15,
            'Kazakhstan': 16, 'Kyrgyzstan': 17, 'Macedonia': 18, 'Malta': 19,
            'Moldova': 20, 'Montenegro': 21, 'Russia': 22, 'Serbia': 23,
            'Turkey': 24, 'Turkmenistan': 25, 'Ukraine': 26, 'Uzbekistan': 27
        }
        for country, row in rest_europe_countries.items():
            index[country] = {
                'Priority': {'sheet': 'Rest of Europe', 'row': row, 'section': 'europe'},
                'Economy': {'sheet': 'Rest of Europe', 'row': row, 'section': 'europe'}
            }
        
        # ROW sheet - Priority section
        row_priority_left = {
            'Canada': 4, 'USA': 5,
            'Antigua & Barbuda': 10, 'Argentina': 11, 'Aruba': 12, 'Bahamas': 13,
            'Barbados': 14, 'Belize': 15, 'Bermuda': 16, 'Bolivia': 17,
            'Brazil': 18, 'Cayman Islands': 19, 'Chile': 20, 'Colombia': 21,
            'Costa Rica': 22, 'Cuba': 23, 'Dominica': 24, 'Dominican Republic': 25,
            'Ecuador': 26, 'El Salvador': 27, 'French Guiana': 28, 'Grenada': 29,
            'Guadeloupe': 30, 'Guatemala': 31, 'Guyana': 32, 'Honduras': 33,
            'Jamaica': 34, 'Martinique': 35, 'Mexico': 36, 'Nicaragua': 37,
            'Panama': 38, 'Paraguay': 39, 'Peru': 40, 'Puerto Rico': 41, 'Reunion': 42
        }
        row_priority_right = {
            'St. Kitts & Nevis': 10, 'St. Lucia': 11, 'St. Vincent': 12,
            'Suriname': 13, 'Trinidad & Tobago': 14, 'Turks & Caicos': 15,
            'Uruguay': 16, 'Venezuela': 17,
            'Bahrain': 22, 'Egypt': 23, 'Iraq': 24, 'Israel': 25, 'Jordan': 26,
            'Kuwait': 27, 'Lebanon': 28, 'Oman': 29, 'Qatar': 30, 'Saudi Arabia': 31, 'UAE': 32,
            'Afghanistan': 37, 'Bangladesh': 38, 'India': 39, 'Nepal': 40,
            'Pakistan': 41, 'Sri Lanka': 42
        }
        
        # ROW sheet - Economy section
        row_economy_left = {
            'Canada': 51, 'USA': 52,
            'Antigua & Barbuda': 57, 'Argentina': 58, 'Aruba': 59, 'Bahamas': 60,
            'Barbados': 61, 'Belize': 62, 'Bermuda': 63, 'Bolivia': 64,
            'Brazil': 65, 'Cayman Islands': 66, 'Chile': 67, 'Colombia': 68,
            'Costa Rica': 69, 'Cuba': 70, 'Dominica': 71, 'Dominican Republic': 72,
            'Ecuador': 73, 'El Salvador': 74, 'French Guiana': 75, 'Grenada': 76,
            'Guadeloupe': 77, 'Guatemala': 78, 'Guyana': 79, 'Honduras': 80,
            'Jamaica': 81, 'Martinique': 82, 'Mexico': 83, 'Nicaragua': 84,
            'Panama': 85, 'Paraguay': 86, 'Peru': 87, 'Puerto Rico': 88, 'Reunion': 89
        }
        row_economy_right = {
            'St. Kitts & Nevis': 57, 'St. Lucia': 58, 'St. Vincent': 59,
            'Suriname': 60, 'Trinidad & Tobago': 61, 'Turks & Caicos': 62,
            'Uruguay': 63, 'Venezuela': 64,
            'Bahrain': 69, 'Egypt': 70, 'Iraq': 71, 'Israel': 72, 'Jordan': 73,
            'Kuwait': 74, 'Lebanon': 75, 'Oman': 76, 'Qatar': 77, 'Saudi Arabia': 78, 'UAE': 79,
            'Afghanistan': 84, 'Bangladesh': 85, 'India': 86, 'Nepal': 87,
            'Pakistan': 88, 'Sri Lanka': 89
        }
        
        for country, row in row_priority_left.items():
            if country not in index:
                index[country] = {}
            index[country]['Priority'] = {'sheet': 'ROW', 'row': row, 'section': 'left'}
        
        for country, row in row_priority_right.items():
            if country not in index:
                index[country] = {}
            index[country]['Priority'] = {'sheet': 'ROW', 'row': row, 'section': 'right'}
        
        for country, row in row_economy_left.items():
            if country not in index:
                index[country] = {}
            index[country]['Economy'] = {'sheet': 'ROW', 'row': row, 'section': 'left'}
        
        for country, row in row_economy_right.items():
            if country not in index:
                index[country] = {}
            index[country]['Economy'] = {'sheet': 'ROW', 'row': row, 'section': 'right'}
        
        # ROW (Continued) - Priority section
        row_cont_priority_left = {
            'Algeria': 4, 'Angola': 5, 'Benin': 6, 'Botswana': 7,
            'Burkina Faso': 8, 'Burundi': 9, 'Cameroon': 10, 'Central African Rep.': 11,
            'Chad': 12, 'Congo, Dem. Rep.': 13, 'Congo, Rep. of': 14, 'Djibouti': 15,
            'Equatorial Guinea': 16, 'Eritrea': 17, 'Ethiopia': 18, 'Gabon': 19,
            'Gambia': 20, 'Ghana': 21, 'Guinea': 22, 'Ivory Coast': 23,
            'Kenya': 24, 'Lesotho': 25, 'Liberia': 26, 'Madagascar': 27,
            'Malawi': 28, 'Maldives': 29, 'Mali': 30, 'Mauritania': 31,
            'Mauritius': 32, 'Morocco': 33, 'Mozambique': 34, 'Namibia': 35,
            'Niger': 36, 'Nigeria': 37, 'Rwanda': 38, 'Senegal': 39,
            'Seychelles': 40, 'Sierra Leone': 41, 'South Africa': 42, 'Sudan': 43,
            'Swaziland': 44, 'Tanzania': 45, 'Togo': 46
        }
        row_cont_priority_right = {
            'Tunisia': 5, 'Uganda': 6, 'Zambia': 12, 'Zimbabwe': 13,
            'American Samoa': 18, 'Australia': 19, 'Brunei Darussalam': 20,
            'Cambodia': 21, 'China': 22, 'Fiji': 23, 'French Polynesia': 24,
            'Hong Kong': 25, 'Indonesia': 26, 'Japan': 27, 'Korea': 28,
            'Laos, Rep. of': 29, 'Malaysia': 30, 'Myanmar': 31, 'New Caledonia': 32,
            'New Zealand': 33, 'Papua New Guinea': 34, 'Philippines': 35,
            'Samoa': 36, 'Singapore': 37, 'Taiwan': 38, 'Thailand': 39,
            'Tonga': 40, 'Vanuatu': 41, 'Vietnam': 42
        }
        
        # ROW (Continued) - Economy section
        row_cont_economy_left = {
            'Algeria': 52, 'Angola': 53, 'Benin': 54, 'Botswana': 55,
            'Burkina Faso': 56, 'Burundi': 57, 'Cameroon': 58, 'Central African Rep.': 59,
            'Chad': 60, 'Congo, Dem. Rep.': 61, 'Congo, Rep. of': 62, 'Djibouti': 63,
            'Equatorial Guinea': 64, 'Eritrea': 65, 'Ethiopia': 66, 'Gabon': 67,
            'Gambia': 68, 'Ghana': 69, 'Guinea': 70, 'Ivory Coast': 71,
            'Kenya': 72, 'Lesotho': 73, 'Liberia': 74, 'Madagascar': 75,
            'Malawi': 76, 'Maldives': 77, 'Mali': 78, 'Mauritania': 79,
            'Mauritius': 80, 'Morocco': 81, 'Mozambique': 82, 'Namibia': 83,
            'Niger': 84, 'Nigeria': 85, 'Rwanda': 86, 'Senegal': 87,
            'Seychelles': 88, 'Sierra Leone': 89, 'South Africa': 90, 'Sudan': 91,
            'Swaziland': 92, 'Tanzania': 93, 'Togo': 94
        }
        row_cont_economy_right = {
            'Tunisia': 53, 'Uganda': 54, 'Zambia': 58, 'Zimbabwe': 59,  # Note: different positions in Economy
            'American Samoa': 62, 'Australia': 63, 'Brunei Darussalam': 64,
            'Cambodia': 65, 'China': 66, 'Fiji': 67, 'French Polynesia': 68,
            'Hong Kong': 69, 'Indonesia': 70, 'Japan': 71, 'Korea': 72,
            'Laos, Rep. of': 73, 'Malaysia': 74, 'Myanmar': 75, 'New Caledonia': 76,
            'New Zealand': 77, 'Papua New Guinea': 78, 'Philippines': 79,
            'Samoa': 80, 'Singapore': 81, 'Taiwan': 82, 'Thailand': 83,
            'Tonga': 84, 'Vanuatu': 85, 'Vietnam': 86
        }
        
        for country, row in row_cont_priority_left.items():
            if country not in index:
                index[country] = {}
            index[country]['Priority'] = {'sheet': 'ROW (Continued)', 'row': row, 'section': 'left'}
        
        for country, row in row_cont_priority_right.items():
            if country not in index:
                index[country] = {}
            index[country]['Priority'] = {'sheet': 'ROW (Continued)', 'row': row, 'section': 'right'}
        
        for country, row in row_cont_economy_left.items():
            if country not in index:
                index[country] = {}
            index[country]['Economy'] = {'sheet': 'ROW (Continued)', 'row': row, 'section': 'left'}
        
        for country, row in row_cont_economy_right.items():
            if country not in index:
                index[country] = {}
            index[country]['Economy'] = {'sheet': 'ROW (Continued)', 'row': row, 'section': 'right'}
        
        return index
    
    def build_country_index(self, workbook) -> Dict[str, dict]:
        """Return the pre-built static index."""
        return self._country_locations
    
    def get_cell_positions(self, country_info: dict, format_type: str) -> Tuple[int, int]:
        """Get columns for items and weight based on sheet and section."""
        section = country_info['section']
        
        if section == 'europe':
            # Determine service from context - we need to get it from the caller
            # For Europe sheets, we use service-based columns
            # This is a limitation - we need to know Priority vs Economy
            # The caller should set this in country_info
            service = country_info.get('service', 'Priority')
            if format_type not in self.EUROPE_COLUMNS[service]:
                raise ValueError(f"Unknown format: {format_type}")
            return self.EUROPE_COLUMNS[service][format_type]
        else:
            # ROW sheets use section-based columns
            if format_type not in self.ROW_COLUMNS[section]:
                raise ValueError(f"Unknown format: {format_type}")
            return self.ROW_COLUMNS[section][format_type]
    
    def set_metadata(self, workbook, po_number: str, shipment_date: str) -> None:
        """Set PO and date in Summary sheet."""
        sheet = workbook['Summary']
        sheet['H7'] = po_number      # Customer Ref (merged cell H7:J7)
        sheet['C8'] = shipment_date  # Shipment Date (merged cell C8:E8)
    
    def place_record(self, workbook, record, country_index: dict):
        """Override to handle Europe service-specific column selection."""
        from .base import PlacementResult
        
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
        
        location = country_data[service].copy()
        location['service'] = service  # Add service for Europe column selection
        
        sheet = workbook[location['sheet']]
        row = location['row']
        
        try:
            items_col, weight_col = self.get_cell_positions(location, format_type)
        except ValueError as e:
            return PlacementResult(
                success=False,
                error_message=str(e)
            )
        
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
        
        sheet.cell(row=row, column=items_col).value = current_items + record.items
        sheet.cell(row=row, column=weight_col).value = round(current_weight + record.weight, 3)
        
        return PlacementResult(
            success=True,
            sheet_name=location['sheet'],
            row=row,
            items_col=items_col,
            weight_col=weight_col
        )
