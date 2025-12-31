"""
Secure credentials management for Multi-Carrier Manifest Tool.

Loads credentials from environment variables or .env file.
NEVER commit credentials to source control.
"""

import os
from dataclasses import dataclass
from typing import Optional
from pathlib import Path


def _load_dotenv():
    """Load .env file if it exists (simple implementation, no dependencies)."""
    # Look for .env in app directory
    app_dir = Path(__file__).parent.parent
    env_file = app_dir / ".env"
    
    if not env_file.exists():
        return
    
    with open(env_file, 'r') as f:
        for line in f:
            line = line.strip()
            # Skip comments and empty lines
            if not line or line.startswith('#'):
                continue
            # Parse KEY=VALUE
            if '=' in line:
                key, _, value = line.partition('=')
                key = key.strip()
                value = value.strip()
                # Remove surrounding quotes if present
                if (value.startswith('"') and value.endswith('"')) or \
                   (value.startswith("'") and value.endswith("'")):
                    value = value[1:-1]
                # Only set if not already in environment
                if key and key not in os.environ:
                    os.environ[key] = value


# Load .env on module import
_load_dotenv()


@dataclass
class PortalCredentials:
    """Credentials for a carrier portal."""
    email: str
    password: str
    contact_name: str = ""
    
    def is_valid(self) -> bool:
        """Check if essential credentials are present."""
        return bool(self.email and self.password)


def get_deutschepost_credentials() -> PortalCredentials:
    """
    Get Deutsche Post portal credentials.
    
    Environment variables:
        DEUTSCHEPOST_EMAIL: Login email
        DEUTSCHEPOST_PASSWORD: Login password
        DEUTSCHEPOST_CONTACT: Contact name for forms (optional)
    """
    return PortalCredentials(
        email=os.environ.get('DEUTSCHEPOST_EMAIL', ''),
        password=os.environ.get('DEUTSCHEPOST_PASSWORD', ''),
        contact_name=os.environ.get('DEUTSCHEPOST_CONTACT', 'Finlay Crawley'),
    )


def get_spring_credentials() -> PortalCredentials:
    """
    Get Spring GDS portal credentials.
    
    Environment variables:
        SPRING_EMAIL: Login email
        SPRING_PASSWORD: Login password
    """
    return PortalCredentials(
        email=os.environ.get('SPRING_EMAIL', ''),
        password=os.environ.get('SPRING_PASSWORD', ''),
    )


def get_landmark_credentials() -> PortalCredentials:
    """
    Get Landmark/bpost portal credentials.
    
    Environment variables:
        LANDMARK_EMAIL: Login email
        LANDMARK_PASSWORD: Login password
    """
    return PortalCredentials(
        email=os.environ.get('LANDMARK_EMAIL', ''),
        password=os.environ.get('LANDMARK_PASSWORD', ''),
    )


def validate_credentials(carrier: str) -> tuple[bool, str]:
    """
    Validate that credentials are configured for a carrier.
    
    Args:
        carrier: One of 'deutschepost', 'spring', 'landmark'
        
    Returns:
        (is_valid, error_message)
    """
    cred_funcs = {
        'deutschepost': get_deutschepost_credentials,
        'spring': get_spring_credentials,
        'landmark': get_landmark_credentials,
    }
    
    if carrier.lower() not in cred_funcs:
        return True, ""  # Unknown carrier, assume no creds needed
    
    creds = cred_funcs[carrier.lower()]()
    
    if not creds.is_valid():
        return False, (
            f"Missing {carrier} portal credentials. "
            f"Set {carrier.upper()}_EMAIL and {carrier.upper()}_PASSWORD "
            f"environment variables or in .env file."
        )
    
    return True, ""
