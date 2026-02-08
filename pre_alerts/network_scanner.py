"""
Network Manifest Scanner
========================
Shared logic for scanning network folders for manifest files.

Used by both the Pre-Alert tab (startup scan) and the
add_manifests_to_queue.py utility script.
"""

import os
import re
from typing import Optional


# Mapping from filename patterns to canonical carrier names
CARRIER_PATTERNS = {
    r"^Air_Business": "Air Business",
    r"^Asendia": "Asendia 2026",
    r"^Deutsche_Post|^Deutsche Post": "Deutsche Post",
    r"^Landmark_Economy": "Landmark Global",
    r"^Landmark_Priority": "Landmark Global",
    r"^PostNord": "PostNord",
    r"^Spring": "Spring",
    r"^United_Business.*NZP": "United Business NZP ETOE",
    r"^United_Business.*SPL": "United Business SPL ETOE",
    r"^United_Business": "United Business ADS",
    r"^Mail_Americas": "Mail Americas",
}


def extract_carrier(filename: str) -> str | None:
    """Extract canonical carrier name from filename."""
    for pattern, carrier in CARRIER_PATTERNS.items():
        if re.match(pattern, filename, re.IGNORECASE):
            return carrier
    return None


def extract_po_number(filename: str) -> str | None:
    """Extract PO number from filename (5-digit number)."""
    match = re.search(r"_(\d{5})_", filename)
    if match:
        return match.group(1)
    return None


def scan_manifests(directory: str, date_filter: str | None = None) -> list[dict]:
    """
    Scan directory for manifest files.

    Args:
        directory: Path to scan
        date_filter: Optional date string (YYYYMMDD) to filter files

    Returns:
        List of dicts with carrier, po_number, path, filename
    """
    manifests = []

    if not os.path.exists(directory):
        return manifests

    for filename in os.listdir(directory):
        # Only process Excel files (manifests)
        if not filename.endswith(".xlsx"):
            continue

        # Skip temp files
        if filename.startswith("~$"):
            continue

        # Optional date filter
        if date_filter and date_filter not in filename:
            continue

        filepath = os.path.join(directory, filename)
        carrier = extract_carrier(filename)
        po_number = extract_po_number(filename)

        if carrier and po_number:
            manifests.append({
                "carrier": carrier,
                "po_number": po_number,
                "path": filepath,
                "filename": filename,
            })

    return manifests


def is_network_path_accessible(path: str) -> bool:
    """
    Check if a network path is accessible.

    Args:
        path: Directory path to check

    Returns:
        True if the path exists and is a directory, False otherwise
    """
    try:
        return os.path.isdir(path)
    except (OSError, PermissionError):
        return False
