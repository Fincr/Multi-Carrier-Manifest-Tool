#!/usr/bin/env python3
"""
Utility script to manually add existing manifests to the pre-alert queue.

This is useful for adding manifests that were created before the queue
integration was enabled, or for re-adding manifests after clearing the queue.

Usage:
    python add_manifests_to_queue.py          # Interactive mode
    python add_manifests_to_queue.py --yes    # Auto-confirm
"""

import argparse
import os
import re
from datetime import datetime
from pre_alerts.manifest_queue import ManifestQueue


# Output directory where manifests are saved
OUTPUT_DIR = r"U:\Erith\Hailey Road\International Ops\Pre-Alerts\Dispatch #1"

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
    # Look for 5-digit PO number pattern
    # Usually appears after carrier name, e.g., "Asendia_2026_27709_..."
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
        List of dicts with carrier, po_number, path
    """
    manifests = []

    if not os.path.exists(directory):
        print(f"Error: Directory not found: {directory}")
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
        else:
            print(f"  Skipped (could not parse): {filename}")

    return manifests


def main():
    parser = argparse.ArgumentParser(
        description="Add existing manifests to the pre-alert queue"
    )
    parser.add_argument(
        "--yes", "-y",
        action="store_true",
        help="Auto-confirm without prompting"
    )
    args = parser.parse_args()

    print("=" * 60)
    print("Add Manifests to Pre-Alert Queue")
    print("=" * 60)
    print()

    # Today's date in filename format
    today = datetime.now().strftime("%Y%m%d")
    print(f"Scanning for today's manifests ({today})...")
    print(f"Directory: {OUTPUT_DIR}")
    print()

    # Scan for manifests
    manifests = scan_manifests(OUTPUT_DIR, date_filter=today)

    if not manifests:
        print("No manifests found for today.")
        return

    print(f"Found {len(manifests)} manifest(s):")
    print("-" * 60)
    for m in manifests:
        print(f"  [{m['carrier']}] PO {m['po_number']}")
        print(f"    {m['filename']}")
    print()

    # Confirm before adding (unless --yes)
    if not args.yes:
        response = input("Add these manifests to the queue? [y/N]: ").strip().lower()
        if response != "y":
            print("Cancelled.")
            return

    # Add to queue
    queue = ManifestQueue()
    added = 0

    print()
    print("Adding to queue...")
    for m in manifests:
        manifest_id = queue.add_manifest(
            carrier=m["carrier"],
            po_number=m["po_number"],
            manifest_path=m["path"],
        )
        print(f"  Added: {m['carrier']} PO {m['po_number']} -> {manifest_id}")
        added += 1

    print()
    print(f"Done! Added {added} manifest(s) to the queue.")
    print()
    print("To verify:")
    print("  1. Launch the main app (python gui.py)")
    print("  2. Go to Pre-Alerts tab")
    print("  3. Manifests should appear in Today's section")


if __name__ == "__main__":
    main()
