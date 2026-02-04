"""
Manifest Queue Persistence Layer
================================
Persistent storage for pre-alert manifests with day-based grouping.
"""

import os
import json
import random
import string
import tempfile
from datetime import datetime, timedelta
from typing import Dict, List, Optional
from dataclasses import dataclass, asdict, field


def get_queue_path() -> str:
    """Get path to manifest_queue.json in the application directory."""
    app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(app_dir, "manifest_queue.json")


@dataclass
class QueuedManifest:
    """A manifest queued for pre-alert sending."""
    id: str
    carrier: str
    po_number: str
    manifest_path: str
    added_at: str  # ISO format: 2026-02-04T14:30:52
    status: str = "pending"  # pending, sent, failed, skipped
    sent_at: Optional[str] = None
    error_message: Optional[str] = None

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> 'QueuedManifest':
        return cls(
            id=data.get('id', ''),
            carrier=data.get('carrier', ''),
            po_number=data.get('po_number', ''),
            manifest_path=data.get('manifest_path', ''),
            added_at=data.get('added_at', ''),
            status=data.get('status', 'pending'),
            sent_at=data.get('sent_at'),
            error_message=data.get('error_message'),
        )

    @property
    def date(self) -> str:
        """Get the date portion of added_at (YYYY-MM-DD)."""
        if self.added_at:
            return self.added_at[:10]
        return datetime.now().strftime("%Y-%m-%d")

    @property
    def time(self) -> str:
        """Get the time portion of added_at (HH:MM)."""
        if self.added_at and len(self.added_at) >= 16:
            return self.added_at[11:16]
        return ""


class ManifestQueue:
    """
    Persistent queue for pre-alert manifests with day-based grouping.

    Data Structure:
    {
        "version": 1,
        "retention_days": 14,
        "manifests": {
            "2026-02-04": [
                {
                    "id": "asendia_2026-02-04_143052_abc1",
                    "carrier": "Asendia",
                    "po_number": "12345",
                    "manifest_path": "C:/path/to/manifest.xlsx",
                    "added_at": "2026-02-04T14:30:52",
                    "status": "pending"
                }
            ]
        }
    }
    """

    VERSION = 1

    def __init__(self, path: Optional[str] = None, retention_days: int = 14):
        self.path = path or get_queue_path()
        self.retention_days = retention_days
        self.data: Dict[str, List[dict]] = {}
        self._load()

    def _load(self):
        """Load queue data from file."""
        if not os.path.exists(self.path):
            self.data = {}
            return

        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                raw = json.load(f)

            # Handle versioned format
            if isinstance(raw, dict) and 'manifests' in raw:
                self.data = raw.get('manifests', {})
                self.retention_days = raw.get('retention_days', self.retention_days)
            else:
                # Legacy format - just the manifests dict
                self.data = raw

        except (json.JSONDecodeError, IOError) as e:
            print(f"Warning: Could not load manifest queue: {e}")
            self.data = {}

    def _save(self):
        """Save queue data to file using atomic write."""
        output = {
            "version": self.VERSION,
            "retention_days": self.retention_days,
            "manifests": self.data
        }

        # Atomic write: write to temp file then rename
        dir_path = os.path.dirname(self.path)
        try:
            with tempfile.NamedTemporaryFile(
                mode='w',
                suffix='.tmp',
                dir=dir_path,
                delete=False,
                encoding='utf-8'
            ) as f:
                json.dump(output, f, indent=2)
                temp_path = f.name

            # On Windows, need to remove target first if it exists
            if os.path.exists(self.path):
                os.remove(self.path)
            os.rename(temp_path, self.path)

        except IOError as e:
            print(f"Warning: Could not save manifest queue: {e}")
            # Clean up temp file if it exists
            if 'temp_path' in locals() and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass

    def _generate_id(self, carrier: str, date: str) -> str:
        """Generate a unique manifest ID."""
        # Slugify carrier name
        carrier_slug = carrier.lower().replace(' ', '_')

        # Time component
        time_str = datetime.now().strftime("%H%M%S")

        # Random suffix
        random_suffix = ''.join(random.choices(string.ascii_lowercase + string.digits, k=4))

        return f"{carrier_slug}_{date}_{time_str}_{random_suffix}"

    # =========================================================================
    # Core CRUD Operations
    # =========================================================================

    def add_manifest(
        self,
        carrier: str,
        po_number: str,
        manifest_path: str,
        date: Optional[str] = None
    ) -> str:
        """
        Add a manifest to the queue.

        Args:
            carrier: Carrier name (canonical)
            po_number: PO number
            manifest_path: Full path to manifest file
            date: Optional date override (YYYY-MM-DD), defaults to today

        Returns:
            The generated manifest ID
        """
        if date is None:
            date = datetime.now().strftime("%Y-%m-%d")

        manifest_id = self._generate_id(carrier, date)
        added_at = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

        manifest = QueuedManifest(
            id=manifest_id,
            carrier=carrier,
            po_number=po_number,
            manifest_path=manifest_path,
            added_at=added_at,
            status="pending"
        )

        if date not in self.data:
            self.data[date] = []

        self.data[date].append(manifest.to_dict())
        self._save()

        return manifest_id

    def get_manifest(self, manifest_id: str) -> Optional[QueuedManifest]:
        """Get a manifest by ID."""
        for date, manifests in self.data.items():
            for m in manifests:
                if m.get('id') == manifest_id:
                    return QueuedManifest.from_dict(m)
        return None

    def update_status(
        self,
        manifest_id: str,
        status: str,
        error_message: Optional[str] = None
    ) -> bool:
        """
        Update the status of a manifest.

        Args:
            manifest_id: The manifest ID
            status: New status (pending, sent, failed, skipped)
            error_message: Optional error message for failed status

        Returns:
            True if found and updated, False otherwise
        """
        for date, manifests in self.data.items():
            for m in manifests:
                if m.get('id') == manifest_id:
                    m['status'] = status
                    if status == 'sent':
                        m['sent_at'] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
                    if error_message:
                        m['error_message'] = error_message
                    self._save()
                    return True
        return False

    def remove_manifest(self, manifest_id: str) -> bool:
        """
        Remove a manifest from the queue.

        Returns:
            True if found and removed, False otherwise
        """
        for date, manifests in self.data.items():
            for i, m in enumerate(manifests):
                if m.get('id') == manifest_id:
                    manifests.pop(i)
                    # Remove empty date entries
                    if not manifests:
                        del self.data[date]
                    self._save()
                    return True
        return False

    # =========================================================================
    # Query Operations
    # =========================================================================

    def get_all_days(self) -> List[str]:
        """
        Get all dates with manifests, sorted newest first.

        Returns:
            List of date strings (YYYY-MM-DD)
        """
        return sorted(self.data.keys(), reverse=True)

    def get_day_manifests(self, date: str) -> List[QueuedManifest]:
        """
        Get all manifests for a specific date.

        Args:
            date: Date string (YYYY-MM-DD)

        Returns:
            List of QueuedManifest objects, sorted by added_at (newest first)
        """
        manifests = self.data.get(date, [])
        result = [QueuedManifest.from_dict(m) for m in manifests]
        # Sort by added_at, newest first
        result.sort(key=lambda m: m.added_at, reverse=True)
        return result

    def get_pending_count(self, date: Optional[str] = None) -> int:
        """
        Get count of pending manifests.

        Args:
            date: If specified, count only for this date. Otherwise count all.

        Returns:
            Count of manifests with status 'pending'
        """
        count = 0
        dates_to_check = [date] if date else self.data.keys()

        for d in dates_to_check:
            if d in self.data:
                for m in self.data[d]:
                    if m.get('status') == 'pending':
                        count += 1

        return count

    def get_sent_count(self, date: Optional[str] = None) -> int:
        """
        Get count of sent manifests.

        Args:
            date: If specified, count only for this date. Otherwise count all.

        Returns:
            Count of manifests with status 'sent'
        """
        count = 0
        dates_to_check = [date] if date else self.data.keys()

        for d in dates_to_check:
            if d in self.data:
                for m in self.data[d]:
                    if m.get('status') == 'sent':
                        count += 1

        return count

    def get_total_count(self) -> int:
        """Get total count of all manifests across all dates."""
        return sum(len(manifests) for manifests in self.data.values())

    def get_day_summary(self, date: str) -> Dict[str, int]:
        """
        Get a summary of statuses for a date.

        Returns:
            Dict with keys: pending, sent, failed, skipped, total
        """
        manifests = self.data.get(date, [])
        summary = {'pending': 0, 'sent': 0, 'failed': 0, 'skipped': 0, 'total': 0}

        for m in manifests:
            status = m.get('status', 'pending')
            if status in summary:
                summary[status] += 1
            summary['total'] += 1

        return summary

    # =========================================================================
    # Maintenance Operations
    # =========================================================================

    def cleanup_old(self, days: Optional[int] = None) -> int:
        """
        Remove manifests older than retention period.

        Args:
            days: Override retention days. Uses instance default if not specified.

        Returns:
            Number of dates removed
        """
        days = days if days is not None else self.retention_days
        cutoff = datetime.now() - timedelta(days=days)
        cutoff_str = cutoff.strftime("%Y-%m-%d")

        old_dates = [d for d in self.data.keys() if d < cutoff_str]
        for date in old_dates:
            del self.data[date]

        if old_dates:
            self._save()

        return len(old_dates)

    def cleanup_missing_files(self) -> int:
        """
        Remove manifests where the file no longer exists.

        Returns:
            Number of manifests removed
        """
        removed = 0
        dates_to_remove = []

        for date, manifests in self.data.items():
            to_remove = []
            for i, m in enumerate(manifests):
                path = m.get('manifest_path', '')
                if path and not os.path.exists(path):
                    to_remove.append(i)
                    removed += 1

            # Remove in reverse order to preserve indices
            for i in reversed(to_remove):
                manifests.pop(i)

            if not manifests:
                dates_to_remove.append(date)

        for date in dates_to_remove:
            del self.data[date]

        if removed > 0:
            self._save()

        return removed

    def clear_all(self):
        """Clear all manifests from the queue."""
        self.data = {}
        self._save()
