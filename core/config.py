"""
Configuration management for Multi-Carrier Manifest Tool.

Handles loading/saving settings to config.json and provides defaults.
"""

import os
import json
from dataclasses import dataclass, asdict
from typing import Optional


# Config file location (same directory as the main script)
def get_config_path() -> str:
    """Get path to config.json in the application directory."""
    app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(app_dir, "config.json")


@dataclass
class AppConfig:
    """Application configuration with defaults."""
    
    # Printer settings
    printer_name: str = "\\\\print01.citipost.co.uk\\KT02"
    
    # Portal settings
    portal_timeout_ms: int = 30000  # 30 seconds
    portal_retry_count: int = 2  # Full workflow retries
    portal_stage_retry_count: int = 2  # Per-stage retries for resilience
    
    # Print settings
    pdf_close_delay_seconds: int = 7
    
    # Processing settings
    max_errors_before_stop: int = 5
    
    # Output settings
    default_output_dir: str = "U:\\Erith\\Hailey Road\\International Ops\\Pre-Alerts\\Dispatch #1"
    
    def save(self, path: Optional[str] = None):
        """Save configuration to JSON file."""
        path = path or get_config_path()
        with open(path, 'w') as f:
            json.dump(asdict(self), f, indent=2)
    
    @classmethod
    def load(cls, path: Optional[str] = None) -> 'AppConfig':
        """
        Load configuration from JSON file.
        Returns default config if file doesn't exist or is invalid.
        """
        path = path or get_config_path()
        
        if not os.path.exists(path):
            return cls()
        
        try:
            with open(path, 'r') as f:
                data = json.load(f)
            
            # Only use known fields, ignore unknown ones
            known_fields = {f.name for f in cls.__dataclass_fields__.values()}
            filtered_data = {k: v for k, v in data.items() if k in known_fields}
            
            return cls(**filtered_data)
        
        except (json.JSONDecodeError, TypeError, KeyError) as e:
            # Invalid config file - return defaults
            print(f"Warning: Could not load config file: {e}")
            return cls()


def get_available_printers() -> list[str]:
    """
    Get list of available Windows printers.
    
    Returns:
        List of printer names, or empty list if enumeration fails.
    """
    printers = []
    
    try:
        import win32print
        
        # Try different enumeration flags
        # PRINTER_ENUM_LOCAL = 2, PRINTER_ENUM_CONNECTIONS = 4
        for flags in [
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS,
            win32print.PRINTER_ENUM_LOCAL,
            win32print.PRINTER_ENUM_CONNECTIONS,
        ]:
            try:
                printer_info = win32print.EnumPrinters(flags, None, 2)
                for printer in printer_info:
                    # printer[2] is the printer name
                    if printer[2] and printer[2] not in printers:
                        printers.append(printer[2])
                if printers:
                    break
            except Exception:
                continue
        
        # If level 2 failed, try level 1 (simpler structure)
        if not printers:
            try:
                printer_info = win32print.EnumPrinters(
                    win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS,
                    None,
                    1
                )
                for printer in printer_info:
                    # Level 1: printer[0] is flags, printer[1] is description, printer[2] is name, printer[3] is comment
                    if len(printer) > 2 and printer[2] and printer[2] not in printers:
                        printers.append(printer[2])
            except Exception:
                pass
        
        # Sort alphabetically for easier selection
        printers.sort()
        
    except ImportError:
        # win32print not available - return empty list
        pass
    except Exception:
        # Silently fail - user can still type printer name manually
        pass
    
    return printers


# Global config instance - loaded once at startup
_config: Optional[AppConfig] = None


def get_config() -> AppConfig:
    """Get the current application configuration."""
    global _config
    if _config is None:
        _config = AppConfig.load()
    return _config


def save_config(config: AppConfig):
    """Save configuration and update global instance."""
    global _config
    config.save()
    _config = config


def reload_config():
    """Force reload configuration from disk."""
    global _config
    _config = AppConfig.load()
