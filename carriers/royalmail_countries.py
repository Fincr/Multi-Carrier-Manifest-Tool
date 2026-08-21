"""
Royal Mail OBA destination countries.

One home for "which destinations can we file, what do carrier sheets call
them, and which OBA region does each sit in". Shared by the carrier sheet
parser (carriers/royalmail.py) and the portal automation
(carriers/royalmail_portal.py) so the two can never disagree about which
countries are supported.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional


# OBA international regions, as they read in the ZZOBA_INTL_REGION dropdown
REGION_EU = "EUROPEAN UNION"
REGION_ROW = "REST OF WORLD"


@dataclass(frozen=True)
class CountrySpec:
    """How one destination country is filed on the OBA order form."""
    region: str
    aliases: List[str] = field(default_factory=list)


@dataclass
class CountryLine:
    """
    One line on the OBA order form: a destination and a format.

    Weight is the sheet's total in KG; the portal wants the average item
    weight in whole grams, which is what avg_weight_grams gives.
    """
    country: str
    format_type: str          # 'Letters' or 'Flats'
    items: int = 0
    weight_kg: float = 0.0

    @property
    def avg_weight_grams(self) -> int:
        if self.items <= 0:
            return 0
        return round(self.weight_kg * 1000 / self.items)


# Insertion order is deliberate: it fixes the order lines appear on the
# OBA order form, so a confirmation PDF always reads the same way.
SUPPORTED_COUNTRIES: Dict[str, CountrySpec] = {
    'Ireland': CountrySpec(
        region=REGION_EU,
        aliases=['Republic of Ireland', 'Eire', 'IE', 'IRL', 'ROI',
                 'Ireland, Republic of', 'Ireland (Republic of)'],
    ),
    'Netherlands': CountrySpec(
        region=REGION_EU,
        aliases=['NL', 'NLD', 'Holland', 'The Netherlands', 'Netherlands (The)'],
    ),
    'Italy': CountrySpec(region=REGION_EU, aliases=['IT', 'ITA']),
    'France': CountrySpec(region=REGION_EU, aliases=['FR', 'FRA']),
    'Germany': CountrySpec(region=REGION_EU, aliases=['DE', 'DEU', 'GER', 'Deutschland']),
    'Canada': CountrySpec(region=REGION_ROW, aliases=['CA', 'CAN']),
}


def normalise_text(value: str) -> str:
    """Collapse whitespace and case-fold, so lookups survive sloppy input."""
    return ' '.join(str(value).split()).lower()


# Alias lookup, built once: every spelling we accept -> canonical name
_ALIAS_LOOKUP: Dict[str, str] = {}
for _name, _spec in SUPPORTED_COUNTRIES.items():
    _ALIAS_LOOKUP[normalise_text(_name)] = _name
    for _alias in _spec.aliases:
        _ALIAS_LOOKUP[normalise_text(_alias)] = _name


def resolve_country(value: Optional[str]) -> Optional[str]:
    """
    Resolve a carrier sheet country value to its canonical name.

    Returns None for anything we cannot file through OBA, so callers can
    decide how loudly to complain.
    """
    if not value:
        return None
    return _ALIAS_LOOKUP.get(normalise_text(value))


def supported_country_list() -> str:
    """Comma-separated canonical names, for error messages."""
    return ', '.join(SUPPORTED_COUNTRIES)
