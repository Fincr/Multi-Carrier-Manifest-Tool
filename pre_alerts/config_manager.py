"""
Pre-Alert Configuration Manager
===============================
Handles loading, saving, and managing pre-alert email configuration.
"""

import os
import json
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional


def get_pre_alert_config_path() -> str:
    """Get path to pre_alert_config.json in the application directory."""
    app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(app_dir, "pre_alert_config.json")


@dataclass
class CarrierEmailConfig:
    """Email configuration for a single carrier."""
    enabled: bool = True
    recipients: List[str] = field(default_factory=list)
    cc: List[str] = field(default_factory=list)
    subject_template: str = "Pre-Alert: {carrier} - {date}"
    
    def to_dict(self) -> dict:
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: dict) -> 'CarrierEmailConfig':
        return cls(
            enabled=data.get('enabled', True),
            recipients=data.get('recipients', []),
            cc=data.get('cc', []),
            subject_template=data.get('subject_template', "Pre-Alert: {carrier} - {date}")
        )


@dataclass 
class PreAlertConfig:
    """Complete pre-alert configuration."""
    
    # Carrier-specific email settings
    carriers: Dict[str, CarrierEmailConfig] = field(default_factory=dict)
    
    # Global settings
    sender_name: str = "Citipost International Operations"
    email_template_path: str = "templates/pre_alert_email.html"
    
    def __post_init__(self):
        """Ensure default carriers are present."""
        default_carriers = {
            "Asendia": CarrierEmailConfig(
                enabled=True,
                recipients=[],
                cc=[],
                subject_template="Pre-Alert: Asendia UK Business - {date}"
            ),
            "Deutsche Post": CarrierEmailConfig(
                enabled=True,
                recipients=[],
                cc=[],
                subject_template="Pre-Alert: Deutsche Post - {date}"
            ),
            "PostNord": CarrierEmailConfig(
                enabled=True,
                recipients=[],
                cc=[],
                subject_template="Pre-Alert: PostNord - {date}"
            ),
            "United Business": CarrierEmailConfig(
                enabled=True,
                recipients=[],
                cc=[],
                subject_template="Pre-Alert: United Business - {date}"
            ),
        }
        
        for carrier, default_config in default_carriers.items():
            if carrier not in self.carriers:
                self.carriers[carrier] = default_config
    
    def get_carrier_config(self, carrier_name: str) -> Optional[CarrierEmailConfig]:
        """
        Get email config for a carrier.
        
        Handles carrier name matching (e.g., "Asendia 2026" -> "Asendia").
        """
        # Direct match
        if carrier_name in self.carriers:
            return self.carriers[carrier_name]
        
        # Partial match (e.g., "Asendia 2026" contains "Asendia")
        carrier_lower = carrier_name.lower()
        for key in self.carriers:
            if key.lower() in carrier_lower or carrier_lower in key.lower():
                return self.carriers[key]
        
        return None
    
    def get_canonical_carrier_name(self, carrier_name: str) -> Optional[str]:
        """
        Get the canonical carrier name for pre-alerts.
        
        Returns the key used in the config, or None if not a pre-alert carrier.
        """
        # Direct match
        if carrier_name in self.carriers:
            return carrier_name
        
        # Partial match
        carrier_lower = carrier_name.lower()
        for key in self.carriers:
            if key.lower() in carrier_lower or carrier_lower in key.lower():
                return key
        
        return None
    
    def is_pre_alert_carrier(self, carrier_name: str) -> bool:
        """Check if a carrier requires pre-alert emails."""
        return self.get_carrier_config(carrier_name) is not None
    
    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            "carriers": {
                name: config.to_dict() 
                for name, config in self.carriers.items()
            },
            "sender_name": self.sender_name,
            "email_template_path": self.email_template_path,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'PreAlertConfig':
        """Create from dictionary (JSON deserialization)."""
        carriers = {}
        if 'carriers' in data:
            for name, config_data in data['carriers'].items():
                carriers[name] = CarrierEmailConfig.from_dict(config_data)
        
        config = cls(
            carriers=carriers,
            sender_name=data.get('sender_name', "Citipost International Operations"),
            email_template_path=data.get('email_template_path', "templates/pre_alert_email.html"),
        )
        
        return config
    
    def save(self, path: Optional[str] = None):
        """Save configuration to JSON file."""
        path = path or get_pre_alert_config_path()
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, indent=2)
    
    @classmethod
    def load(cls, path: Optional[str] = None) -> 'PreAlertConfig':
        """
        Load configuration from JSON file.
        Returns default config if file doesn't exist or is invalid.
        """
        path = path or get_pre_alert_config_path()
        
        if not os.path.exists(path):
            return cls()
        
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return cls.from_dict(data)
        
        except (json.JSONDecodeError, TypeError, KeyError) as e:
            print(f"Warning: Could not load pre-alert config: {e}")
            return cls()


# Convenience functions
def load_pre_alert_config() -> PreAlertConfig:
    """Load the pre-alert configuration."""
    return PreAlertConfig.load()


def save_pre_alert_config(config: PreAlertConfig):
    """Save the pre-alert configuration."""
    config.save()
