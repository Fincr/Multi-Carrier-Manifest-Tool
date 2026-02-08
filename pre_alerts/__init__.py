"""
Pre-Alert Email Module
======================
Handles automated pre-alert emails for carrier manifests.

Supports:
- Asendia
- Deutsche Post
- PostNord
- United Business

Uses Outlook COM automation to send emails from the user's desktop client.
"""

from .config_manager import PreAlertConfig, QueueSettings, load_pre_alert_config, save_pre_alert_config
from .email_sender import send_pre_alert_email, OutlookEmailSender
from .send_tracker import SendTracker
from .manifest_queue import ManifestQueue, QueuedManifest
from .network_scanner import (
    CARRIER_PATTERNS, extract_carrier, extract_po_number,
    scan_manifests, is_network_path_accessible,
)

__all__ = [
    'PreAlertConfig',
    'QueueSettings',
    'load_pre_alert_config',
    'save_pre_alert_config',
    'send_pre_alert_email',
    'OutlookEmailSender',
    'SendTracker',
    'ManifestQueue',
    'QueuedManifest',
    'CARRIER_PATTERNS',
    'extract_carrier',
    'extract_po_number',
    'scan_manifests',
    'is_network_path_accessible',
]
