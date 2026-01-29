"""
Carrier modules for manifest population.
"""

from .base import BaseCarrier, ShipmentRecord as ShipmentRecord, PlacementResult as PlacementResult
from .asendia import AsendiaCarrier, Asendia2025Carrier
from .postnord import PostNordCarrier
from .spring import SpringCarrier
from .airbusiness import AirBusinessCarrier
from .mail_americas import MailAmericasCarrier
from .landmark import LandmarkCarrier
from .deutschepost import DeutschePostCarrier
from .unitedbusiness import UnitedBusinessCarrier
from .unitedbusiness_nzp import UnitedBusinessNZPCarrier

# Registry of available carriers
CARRIER_REGISTRY = {
    'Asendia 2026': AsendiaCarrier,
    'Asendia 2025': Asendia2025Carrier,
    'PostNord': PostNordCarrier,
    'Spring': SpringCarrier,
    'Air Business': AirBusinessCarrier,
    'Mail Americas': MailAmericasCarrier,
    'Landmark Global': LandmarkCarrier,
    'Deutsche Post': DeutschePostCarrier,
    'United Business ADS': UnitedBusinessCarrier,
    'United Business NZP ETOE': UnitedBusinessNZPCarrier,
}


def get_carrier(carrier_name: str) -> BaseCarrier:
    """Get carrier handler by name."""
    # Try exact match first
    if carrier_name in CARRIER_REGISTRY:
        return CARRIER_REGISTRY[carrier_name]()

    # Normalise: lowercase, strip whitespace
    carrier_lower = carrier_name.lower().strip()

    # Explicitly exclude unsupported carriers
    if 'jersey post' in carrier_lower:
        raise ValueError(f"Jersey Post is not a supported carrier: {carrier_name}")

    if 'publications' in carrier_lower:
        raise ValueError(f"Asendia Publications is not a supported carrier: {carrier_name}")
    
    if 'mmp parcel' in carrier_lower:
        raise ValueError(f"PostNord MMP Parcel is not a supported carrier (use standard PostNord): {carrier_name}")
    
    if 'lettershop' in carrier_lower:
        raise ValueError(f"Lettershop is not a supported carrier: {carrier_name}")
    
    # Check for year-specific Asendia matches
    if 'asendia' in carrier_lower:
        if '2025' in carrier_name:
            return CARRIER_REGISTRY['Asendia 2025']()
        else:
            # Default to 2026 for any other Asendia variant
            return CARRIER_REGISTRY['Asendia 2026']()
    
    # Check for PostNord (ignore any year suffix)
    if 'postnord' in carrier_lower:
        return CARRIER_REGISTRY['PostNord']()
    
    # Check for Spring
    if 'spring' in carrier_lower:
        return CARRIER_REGISTRY['Spring']()
    
    # Check for Air Business
    if 'air business' in carrier_lower or 'airbusiness' in carrier_lower:
        return CARRIER_REGISTRY['Air Business']()
    
    # Check for Mail Americas (matches "Mail Americas", "Mail Americas Non Ready", etc.)
    if 'mail americas' in carrier_lower or 'mail africa' in carrier_lower:
        return CARRIER_REGISTRY['Mail Americas']()
    
    # Check for Landmark Global
    if 'landmark' in carrier_lower:
        return CARRIER_REGISTRY['Landmark Global']()
    
    # Check for Deutsche Post
    if 'deutsche' in carrier_lower or 'deutschepost' in carrier_lower:
        return CARRIER_REGISTRY['Deutsche Post']()
    
    # Check for United Business variants
    if 'united business' in carrier_lower or 'ubl' in carrier_lower:
        # Check for NZP ETOE variant first (more specific)
        if 'nzp' in carrier_lower or 'etoe' in carrier_lower or 't&d' in carrier_lower or 't d' in carrier_lower:
            return CARRIER_REGISTRY['United Business NZP ETOE']()
        # Default to ADS variant
        return CARRIER_REGISTRY['United Business ADS']()
    
    raise ValueError(f"Unknown carrier: {carrier_name}. Available: {list(CARRIER_REGISTRY.keys())}")


def list_carriers() -> list:
    """List available carrier names."""
    return list(CARRIER_REGISTRY.keys())
