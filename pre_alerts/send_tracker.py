"""
Pre-Alert Send Tracker
======================
Tracks which pre-alerts have been sent to prevent duplicates.
"""

import os
import json
from datetime import datetime
from typing import Dict, Optional
from dataclasses import dataclass, asdict


def get_tracker_path() -> str:
    """Get path to pre_alert_log.json in the application directory."""
    app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(app_dir, "pre_alert_log.json")


@dataclass
class SendRecord:
    """Record of a sent pre-alert."""
    po_number: str
    manifest_path: str
    sent_at: str  # ISO format timestamp
    recipients: list
    cc: list
    success: bool
    error_message: str = ""
    
    def to_dict(self) -> dict:
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: dict) -> 'SendRecord':
        return cls(
            po_number=data.get('po_number', ''),
            manifest_path=data.get('manifest_path', ''),
            sent_at=data.get('sent_at', ''),
            recipients=data.get('recipients', []),
            cc=data.get('cc', []),
            success=data.get('success', False),
            error_message=data.get('error_message', ''),
        )


class SendTracker:
    """
    Tracks sent pre-alerts by date and carrier.
    
    Structure:
    {
        "2026-01-23": {
            "Asendia": {
                "po_number": "12345",
                "manifest_path": "...",
                "sent_at": "14:32:05",
                "recipients": [...],
                "cc": [...],
                "success": true
            },
            "PostNord": {...}
        },
        "2026-01-24": {...}
    }
    """
    
    def __init__(self, path: Optional[str] = None):
        self.path = path or get_tracker_path()
        self.data: Dict[str, Dict[str, dict]] = {}
        self._load()
    
    def _load(self):
        """Load tracker data from file."""
        if not os.path.exists(self.path):
            self.data = {}
            return
        
        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            print(f"Warning: Could not load send tracker: {e}")
            self.data = {}
    
    def _save(self):
        """Save tracker data to file."""
        try:
            with open(self.path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, indent=2)
        except IOError as e:
            print(f"Warning: Could not save send tracker: {e}")
    
    def _get_today(self) -> str:
        """Get today's date as string."""
        return datetime.now().strftime("%Y-%m-%d")
    
    def was_sent_today(self, carrier_name: str) -> bool:
        """Check if a pre-alert was already sent today for this carrier."""
        today = self._get_today()
        if today not in self.data:
            return False
        return carrier_name in self.data[today]
    
    def get_today_record(self, carrier_name: str) -> Optional[SendRecord]:
        """Get the send record for a carrier from today, if exists."""
        today = self._get_today()
        if today not in self.data:
            return None
        if carrier_name not in self.data[today]:
            return None
        return SendRecord.from_dict(self.data[today][carrier_name])
    
    def record_send(self, carrier_name: str, record: SendRecord):
        """Record that a pre-alert was sent."""
        today = self._get_today()
        if today not in self.data:
            self.data[today] = {}
        
        self.data[today][carrier_name] = record.to_dict()
        self._save()
    
    def get_all_today(self) -> Dict[str, SendRecord]:
        """Get all send records from today."""
        today = self._get_today()
        if today not in self.data:
            return {}
        
        return {
            carrier: SendRecord.from_dict(data)
            for carrier, data in self.data[today].items()
        }
    
    def clear_today(self, carrier_name: Optional[str] = None):
        """
        Clear today's records.
        
        Args:
            carrier_name: If specified, only clear this carrier. Otherwise clear all.
        """
        today = self._get_today()
        if today not in self.data:
            return
        
        if carrier_name:
            if carrier_name in self.data[today]:
                del self.data[today][carrier_name]
        else:
            del self.data[today]
        
        self._save()
    
    def cleanup_old_records(self, days_to_keep: int = 30):
        """Remove records older than specified days."""
        from datetime import timedelta
        
        cutoff = datetime.now() - timedelta(days=days_to_keep)
        cutoff_str = cutoff.strftime("%Y-%m-%d")
        
        old_dates = [d for d in self.data.keys() if d < cutoff_str]
        for date in old_dates:
            del self.data[date]
        
        if old_dates:
            self._save()
