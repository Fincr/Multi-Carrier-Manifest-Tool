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
from datetime import datetime
from pre_alerts.manifest_queue import ManifestQueue
from pre_alerts.network_scanner import (
    CARRIER_PATTERNS, extract_carrier, extract_po_number, scan_manifests,
)


# Output directory where manifests are saved
OUTPUT_DIR = r"U:\Erith\Hailey Road\International Ops\Pre-Alerts\Dispatch #1"


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
            date=m.get("date"),
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
