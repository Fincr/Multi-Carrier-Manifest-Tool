"""
Pre-Alert Tab UI
================
Tkinter tab for managing and sending pre-alert emails.
"""

import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Callable
import threading

from .config_manager import (
    PreAlertConfig, CarrierEmailConfig, QueueSettings,
    load_pre_alert_config, save_pre_alert_config
)
from .send_tracker import SendTracker, SendRecord
from .manifest_queue import ManifestQueue, QueuedManifest
from .email_sender import send_pre_alert_email
from .network_scanner import scan_manifests, is_network_path_accessible
from core.config import get_config


class PreAlertConfigDialog:
    """Dialog for editing pre-alert email configuration."""

    def __init__(self, parent, config: PreAlertConfig, carrier_name: str):
        self.result = None
        self.config = config
        self.carrier_name = carrier_name
        self.carrier_config = config.carriers.get(carrier_name, CarrierEmailConfig())

        # Create dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Configure {carrier_name} Pre-Alert")
        self.dialog.geometry("500x400")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (500 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (400 // 2)
        self.dialog.geometry(f"+{x}+{y}")

        self._create_widgets()

    def _create_widgets(self):
        """Build the dialog UI."""
        main_frame = ttk.Frame(self.dialog, padding="15")
        main_frame.pack(fill='both', expand=True)

        # Enabled checkbox
        self.enabled_var = tk.BooleanVar(value=self.carrier_config.enabled)
        ttk.Checkbutton(
            main_frame,
            text="Enable pre-alerts for this carrier",
            variable=self.enabled_var
        ).pack(anchor='w', pady=(0, 10))

        # Recipients
        ttk.Label(main_frame, text="To (one email per line):").pack(anchor='w')
        self.recipients_text = scrolledtext.ScrolledText(main_frame, height=4, width=50)
        self.recipients_text.pack(fill='x', pady=(0, 10))
        self.recipients_text.insert('1.0', '\n'.join(self.carrier_config.recipients))

        # CC
        ttk.Label(main_frame, text="CC (one email per line):").pack(anchor='w')
        self.cc_text = scrolledtext.ScrolledText(main_frame, height=3, width=50)
        self.cc_text.pack(fill='x', pady=(0, 10))
        self.cc_text.insert('1.0', '\n'.join(self.carrier_config.cc))

        # Subject template
        ttk.Label(main_frame, text="Subject template:").pack(anchor='w')
        self.subject_var = tk.StringVar(value=self.carrier_config.subject_template)
        subject_entry = ttk.Entry(main_frame, textvariable=self.subject_var, width=60)
        subject_entry.pack(fill='x', pady=(0, 5))

        ttk.Label(
            main_frame,
            text="Available placeholders: {carrier}, {date}, {po_number}",
            font=('Helvetica', 8),
            foreground='gray'
        ).pack(anchor='w', pady=(0, 15))

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x')

        ttk.Button(button_frame, text="Save", command=self._save).pack(side='right', padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self._cancel).pack(side='right')

    def _parse_emails(self, text: str) -> List[str]:
        """Parse email addresses from text (one per line)."""
        lines = text.strip().split('\n')
        emails = []
        for line in lines:
            email = line.strip()
            if email and '@' in email:
                emails.append(email)
        return emails

    def _save(self):
        """Save configuration and close."""
        recipients = self._parse_emails(self.recipients_text.get('1.0', 'end'))
        cc = self._parse_emails(self.cc_text.get('1.0', 'end'))
        subject = self.subject_var.get().strip()

        if self.enabled_var.get() and not recipients:
            messagebox.showerror("Error", "Please enter at least one recipient email address.")
            return

        if not subject:
            messagebox.showerror("Error", "Please enter a subject template.")
            return

        # Update config
        self.carrier_config.enabled = self.enabled_var.get()
        self.carrier_config.recipients = recipients
        self.carrier_config.cc = cc
        self.carrier_config.subject_template = subject

        self.config.carriers[self.carrier_name] = self.carrier_config
        save_pre_alert_config(self.config)

        self.result = self.config
        self.dialog.destroy()

    def _cancel(self):
        """Close without saving."""
        self.result = None
        self.dialog.destroy()

    def show(self) -> Optional[PreAlertConfig]:
        """Show dialog and return result."""
        self.dialog.wait_window()
        return self.result


class PreAlertTab:
    """
    Pre-Alert management tab for the manifest tool.

    Displays processed manifests that require pre-alerts and allows
    the user to select which ones to send.
    """

    # Carriers that require pre-alerts
    PRE_ALERT_CARRIERS = ["Asendia", "Deutsche Post", "PostNord", "United Business"]

    def __init__(self, parent: ttk.Frame, app_dir: str):
        """
        Initialize the pre-alert tab.

        Args:
            parent: Parent frame (notebook tab)
            app_dir: Application directory for locating config/templates
        """
        self.parent = parent
        self.app_dir = app_dir
        self.config = load_pre_alert_config()
        self.tracker = SendTracker()
        self.queue = ManifestQueue(retention_days=self.config.queue_settings.retention_days)

        # Cleanup on startup if enabled
        if self.config.queue_settings.auto_cleanup:
            self.queue.cleanup_old()
            self.queue.cleanup_missing_files()

        # Log callback (set by main app)
        self.log_callback: Optional[Callable[[str], None]] = None

        self._create_widgets()
        self._refresh_display()
        self._start_network_scan()

    def set_log_callback(self, callback: Callable[[str], None]):
        """Set callback for logging messages."""
        self.log_callback = callback

    def log(self, message: str):
        """Log a message."""
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)

    def _create_widgets(self):
        """Build the tab UI."""
        # Main container
        main_frame = ttk.Frame(self.parent, padding="10")
        main_frame.pack(fill='both', expand=True)

        # Title and description
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill='x', pady=(0, 10))

        ttk.Label(
            title_frame,
            text="Pre-Alert Emails",
            font=('Helvetica', 14, 'bold')
        ).pack(side='left')

        ttk.Button(
            title_frame,
            text="Settings",
            command=self._open_global_settings
        ).pack(side='right')

        ttk.Label(
            main_frame,
            text="Select manifests to send pre-alert emails. Only carriers requiring pre-alerts are shown.",
            foreground='gray'
        ).pack(anchor='w', pady=(0, 10))

        # Summary panel
        self.summary_frame = ttk.Frame(main_frame)
        self.summary_frame.pack(fill='x', pady=(0, 10))

        self.summary_label = ttk.Label(
            self.summary_frame,
            text="Queue: 0 pending | 0 sent today | 0 total",
            font=('Helvetica', 10)
        )
        self.summary_label.pack(side='left')

        # Manifest list frame
        list_frame = ttk.LabelFrame(main_frame, text="Available Manifests", padding="10")
        list_frame.pack(fill='both', expand=True, pady=(0, 10))

        # Create hierarchical treeview for manifest list
        columns = ('carrier', 'po', 'file', 'status', 'time')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=12)

        # Date column (tree column)
        self.tree.heading('#0', text='Date')
        self.tree.column('#0', width=180, stretch=False)

        # Data columns
        self.tree.heading('carrier', text='Carrier')
        self.tree.heading('po', text='PO Number')
        self.tree.heading('file', text='Manifest File')
        self.tree.heading('status', text='Status')
        self.tree.heading('time', text='Time')

        self.tree.column('carrier', width=120)
        self.tree.column('po', width=80)
        self.tree.column('file', width=220)
        self.tree.column('status', width=100)
        self.tree.column('time', width=60)

        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Bind right-click for context menu
        self.tree.bind('<Button-3>', self._show_context_menu)

        # Create context menu
        self._create_context_menu()

        # Selection buttons
        select_frame = ttk.Frame(main_frame)
        select_frame.pack(fill='x', pady=(0, 10))

        ttk.Button(select_frame, text="Select All", command=self._select_all).pack(side='left', padx=(0, 5))
        ttk.Button(select_frame, text="Select None", command=self._select_none).pack(side='left', padx=(0, 5))
        ttk.Button(select_frame, text="Select Day", command=self._select_day).pack(side='left', padx=(0, 5))
        ttk.Button(select_frame, text="Select Pending", command=self._select_pending).pack(side='left', padx=(0, 5))
        ttk.Button(select_frame, text="Configure Selected", command=self._configure_selected).pack(side='left', padx=(0, 5))
        ttk.Button(select_frame, text="Refresh", command=self._refresh_display).pack(side='right')

        # Action buttons
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill='x', pady=(0, 10))

        self.send_button = ttk.Button(
            action_frame,
            text="Send Pre-Alerts",
            command=self._send_selected
        )
        self.send_button.pack(side='left', padx=(0, 10))

        ttk.Button(
            action_frame,
            text="Mark Skipped",
            command=self._mark_skipped
        ).pack(side='left', padx=(0, 5))

        ttk.Button(
            action_frame,
            text="Remove Selected",
            command=self._remove_selected
        ).pack(side='left', padx=(0, 5))

        ttk.Button(
            action_frame,
            text="Purge Old",
            command=self._purge_old
        ).pack(side='right')

        # Log area
        log_frame = ttk.LabelFrame(main_frame, text="Send Log", padding="5")
        log_frame.pack(fill='both', expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, font=('Consolas', 9))
        self.log_text.pack(fill='both', expand=True)

    def _create_context_menu(self):
        """Create the right-click context menu."""
        self.context_menu = tk.Menu(self.parent, tearoff=0)
        self.context_menu.add_command(label="Send Pre-Alert", command=self._send_selected)
        self.context_menu.add_command(label="Configure Carrier", command=self._configure_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Mark as Skipped", command=self._mark_skipped)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Open Manifest File", command=self._open_manifest_file)
        self.context_menu.add_command(label="Open Containing Folder", command=self._open_containing_folder)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Remove from Queue", command=self._remove_selected)

    def _show_context_menu(self, event):
        """Show context menu on right-click."""
        # Select item under cursor
        item = self.tree.identify_row(event.y)
        if item:
            # If clicked item is not in selection, select only it
            if item not in self.tree.selection():
                self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def _format_date_display(self, date_str: str) -> str:
        """Format date for display with friendly labels."""
        try:
            date = datetime.strptime(date_str, "%Y-%m-%d")
            today = datetime.now().date()
            date_date = date.date()

            if date_date == today:
                return f"Today ({date.strftime('%b %d')})"
            elif date_date == today - timedelta(days=1):
                return f"Yesterday ({date.strftime('%b %d')})"
            elif (today - date_date).days < 7:
                return date.strftime("%a (%b %d)")
            else:
                return date_str
        except ValueError:
            return date_str

    def _refresh_display(self):
        """Refresh the manifest list display with hierarchical day grouping."""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Get all days from queue
        days = self.queue.get_all_days()
        today = datetime.now().strftime("%Y-%m-%d")

        for date in days:
            manifests = self.queue.get_day_manifests(date)
            if not manifests:
                continue

            summary = self.queue.get_day_summary(date)
            date_display = self._format_date_display(date)

            # Build summary text
            summary_parts = []
            if summary['pending'] > 0:
                summary_parts.append(f"{summary['pending']} pending")
            if summary['sent'] > 0:
                summary_parts.append(f"{summary['sent']} sent")
            if summary['failed'] > 0:
                summary_parts.append(f"{summary['failed']} failed")
            if summary['skipped'] > 0:
                summary_parts.append(f"{summary['skipped']} skipped")

            summary_text = " - " + ", ".join(summary_parts) if summary_parts else ""

            # Insert date group parent node
            date_iid = f"date_{date}"
            self.tree.insert(
                '',
                'end',
                iid=date_iid,
                text=f"{date_display}{summary_text}",
                open=(date == today and self.config.queue_settings.expand_today_on_start)
            )

            # Insert child manifest nodes
            for manifest in manifests:
                # Determine status display
                status_display = self._get_status_display(manifest)

                self.tree.insert(
                    date_iid,
                    'end',
                    iid=manifest.id,
                    values=(
                        manifest.carrier,
                        manifest.po_number,
                        os.path.basename(manifest.manifest_path),
                        status_display,
                        manifest.time
                    )
                )

        # Update summary panel
        self._update_summary()

    def _get_status_display(self, manifest: QueuedManifest) -> str:
        """Get display text for manifest status."""
        if manifest.status == 'sent':
            sent_time = ""
            if manifest.sent_at and len(manifest.sent_at) >= 16:
                sent_time = manifest.sent_at[11:16]
            return f"Sent {sent_time}"
        elif manifest.status == 'failed':
            return "Failed"
        elif manifest.status == 'skipped':
            return "Skipped"
        else:
            # Check if carrier is configured
            carrier_config = self.config.get_carrier_config(manifest.carrier)
            if carrier_config and carrier_config.enabled and carrier_config.recipients:
                return "Ready"
            elif carrier_config and not carrier_config.enabled:
                return "Disabled"
            else:
                return "Not configured"

    def _update_summary(self):
        """Update the summary panel."""
        today = datetime.now().strftime("%Y-%m-%d")
        pending = self.queue.get_pending_count()
        sent_today = self.queue.get_sent_count(today)
        total = self.queue.get_total_count()

        self.summary_label.config(
            text=f"Queue: {pending} pending | {sent_today} sent today | {total} total"
        )

    def _get_selected_manifests(self) -> List[QueuedManifest]:
        """Get manifest objects for selected items, excluding date groups."""
        selected = self.tree.selection()
        manifests = []

        for iid in selected:
            # Skip date group items
            if iid.startswith('date_'):
                continue
            manifest = self.queue.get_manifest(iid)
            if manifest:
                manifests.append(manifest)

        return manifests

    def _select_all(self):
        """Select all manifest items (not date groups)."""
        for date_item in self.tree.get_children():
            for manifest_item in self.tree.get_children(date_item):
                self.tree.selection_add(manifest_item)

    def _select_none(self):
        """Deselect all items."""
        self.tree.selection_remove(*self.tree.selection())

    def _select_day(self):
        """Select all manifests for the currently selected day."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a date or manifest first.")
            return

        # Get the date from selection
        first = selected[0]
        if first.startswith('date_'):
            date_item = first
        else:
            date_item = self.tree.parent(first)

        if date_item:
            for manifest_item in self.tree.get_children(date_item):
                self.tree.selection_add(manifest_item)

    def _select_pending(self):
        """Select all manifests with pending status."""
        for date_item in self.tree.get_children():
            for manifest_item in self.tree.get_children(date_item):
                manifest = self.queue.get_manifest(manifest_item)
                if manifest and manifest.status == 'pending':
                    self.tree.selection_add(manifest_item)

    def _configure_selected(self):
        """Open configuration dialog for selected carrier."""
        manifests = self._get_selected_manifests()
        if not manifests:
            messagebox.showwarning("No Selection", "Please select a manifest to configure.")
            return

        # Get first selected carrier
        carrier = manifests[0].carrier

        # Open config dialog
        dialog = PreAlertConfigDialog(self.parent, self.config, carrier)
        result = dialog.show()

        if result:
            self.config = result
            self._refresh_display()

    def _open_global_settings(self):
        """Open global pre-alert settings dialog."""
        dialog = GlobalSettingsDialog(self.parent, self.config)
        result = dialog.show()

        if result:
            self.config = result
            self._refresh_display()

    def _send_selected(self):
        """Send pre-alerts for selected manifests."""
        manifests = self._get_selected_manifests()
        if not manifests:
            messagebox.showwarning("No Selection", "Please select at least one manifest to send.")
            return

        # Validate selected manifests
        to_send = []
        for manifest in manifests:
            carrier_config = self.config.get_carrier_config(manifest.carrier)

            if not carrier_config or not carrier_config.enabled:
                self._log(f"[!] {manifest.carrier}: Skipped (disabled)")
                continue

            if not carrier_config.recipients:
                self._log(f"[!] {manifest.carrier}: Skipped (no recipients configured)")
                continue

            # Warn if already sent
            if manifest.status == 'sent':
                if not messagebox.askyesno(
                    "Already Sent",
                    f"Pre-alert for {manifest.carrier} (PO {manifest.po_number}) was already sent.\n\nSend again?"
                ):
                    self._log(f"[!] {manifest.carrier}: Skipped (already sent)")
                    continue

            to_send.append((manifest, carrier_config))

        if not to_send:
            messagebox.showinfo("Nothing to Send", "No valid manifests to send.")
            return

        # Confirm
        msg = f"Send pre-alerts for {len(to_send)} manifest(s)?\n\n"
        for manifest, _ in to_send:
            msg += f"  {manifest.carrier} - PO {manifest.po_number}\n"

        if not messagebox.askyesno("Confirm Send", msg):
            return

        # Disable send button during processing
        self.send_button.config(state='disabled')

        # Send in background thread
        thread = threading.Thread(
            target=self._send_emails_thread,
            args=(to_send,),
            daemon=True
        )
        thread.start()

    def _send_emails_thread(self, to_send: list):
        """Background thread for sending emails."""
        results = []

        try:
            for manifest, carrier_config in to_send:
                self.parent.after(0, self._log, f"\nSending pre-alert for {manifest.carrier}...")

                try:
                    # Get template path
                    template_path = os.path.join(self.app_dir, self.config.email_template_path)

                    # Send email
                    success, message = send_pre_alert_email(
                        carrier_name=manifest.carrier,
                        po_number=manifest.po_number,
                        manifest_path=manifest.manifest_path,
                        recipients=carrier_config.recipients,
                        cc=carrier_config.cc,
                        subject_template=carrier_config.subject_template,
                        sender_name=self.config.sender_name,
                        template_path=template_path
                    )
                except Exception as e:
                    success = False
                    message = f"Error: {str(e)}"

                # Record result in send tracker (for compatibility)
                record = SendRecord(
                    po_number=manifest.po_number,
                    manifest_path=manifest.manifest_path,
                    sent_at=datetime.now().strftime("%H:%M:%S"),
                    recipients=carrier_config.recipients,
                    cc=carrier_config.cc,
                    success=success,
                    error_message="" if success else message
                )
                self.tracker.record_send(manifest.carrier, record)

                # Update queue status
                if success:
                    self.queue.update_status(manifest.id, 'sent')
                    self.parent.after(0, self._log, f"  [OK] {message}")
                    self.parent.after(0, self._log, f"    To: {', '.join(carrier_config.recipients)}")
                    if carrier_config.cc:
                        self.parent.after(0, self._log, f"    CC: {', '.join(carrier_config.cc)}")
                else:
                    self.queue.update_status(manifest.id, 'failed', message)
                    self.parent.after(0, self._log, f"  [FAIL] {message}")

                results.append((manifest.carrier, success, message))
        except Exception as e:
            self.parent.after(0, self._log, f"\n[ERROR] Thread crashed: {e}")
        finally:
            # Always call completion handler to re-enable UI
            self.parent.after(0, self._on_send_complete, results)

    def _on_send_complete(self, results: list):
        """Handle send completion."""
        self.send_button.config(state='normal')
        self._refresh_display()

        # Summary
        successful = sum(1 for _, success, _ in results if success)
        failed = len(results) - successful

        self._log(f"\n{'='*40}")
        self._log(f"COMPLETE: {successful} sent, {failed} failed")

        if failed > 0:
            messagebox.showwarning(
                "Partially Complete",
                f"Pre-alerts sent: {successful}\nFailed: {failed}\n\nCheck the log for details."
            )
        else:
            messagebox.showinfo(
                "Complete",
                f"Successfully sent {successful} pre-alert(s)."
            )

    def _mark_skipped(self):
        """Mark selected manifests as skipped."""
        manifests = self._get_selected_manifests()
        if not manifests:
            messagebox.showwarning("No Selection", "Please select manifests to mark as skipped.")
            return

        for manifest in manifests:
            self.queue.update_status(manifest.id, 'skipped')

        self._refresh_display()
        self._log(f"Marked {len(manifests)} manifest(s) as skipped")

    def _remove_selected(self):
        """Remove selected manifests from queue."""
        manifests = self._get_selected_manifests()
        if not manifests:
            messagebox.showwarning("No Selection", "Please select manifests to remove.")
            return

        if not messagebox.askyesno(
            "Confirm Remove",
            f"Remove {len(manifests)} manifest(s) from the queue?\n\nThis cannot be undone."
        ):
            return

        for manifest in manifests:
            self.queue.remove_manifest(manifest.id)

        self._refresh_display()
        self._log(f"Removed {len(manifests)} manifest(s) from queue")

    def _purge_old(self):
        """Manually trigger cleanup of old entries."""
        days = self.config.queue_settings.retention_days

        if not messagebox.askyesno(
            "Confirm Purge",
            f"Remove all manifests older than {days} days?"
        ):
            return

        removed = self.queue.cleanup_old(days)
        missing = self.queue.cleanup_missing_files()

        self._refresh_display()
        self._log(f"Purged {removed} old date(s), {missing} missing file(s)")

    def _open_manifest_file(self):
        """Open the selected manifest file."""
        manifests = self._get_selected_manifests()
        if not manifests:
            return

        path = manifests[0].manifest_path
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showerror("File Not Found", f"Manifest file not found:\n{path}")

    def _open_containing_folder(self):
        """Open the folder containing the selected manifest."""
        manifests = self._get_selected_manifests()
        if not manifests:
            return

        path = manifests[0].manifest_path
        folder = os.path.dirname(path)
        if os.path.exists(folder):
            subprocess.run(['explorer', folder])
        else:
            messagebox.showerror("Folder Not Found", f"Folder not found:\n{folder}")

    def _log(self, message: str):
        """Add message to the log display."""
        self.log_text.insert('end', message + '\n')
        self.log_text.see('end')

        # Also send to main app log if callback set
        if self.log_callback:
            self.log_callback(message)

    def add_manifest(self, carrier_name: str, po_number: str, manifest_path: str):
        """
        Add a processed manifest to the pre-alert queue.

        Called by the main app after processing a manifest.

        Args:
            carrier_name: Name of the carrier (as detected from sheet)
            po_number: PO number
            manifest_path: Full path to the output manifest file
        """
        # Check if this is a pre-alert carrier
        canonical = self.config.get_canonical_carrier_name(carrier_name)
        if not canonical:
            return  # Not a pre-alert carrier

        # Add to persistent queue
        manifest_id = self.queue.add_manifest(
            carrier=canonical,
            po_number=po_number,
            manifest_path=manifest_path
        )

        self._log(f"Queued: {canonical} - PO {po_number}")
        self._refresh_display()

    def clear_manifests(self):
        """Clear all manifests from the queue."""
        if messagebox.askyesno(
            "Confirm Clear",
            "Clear all manifests from the queue?\n\nThis cannot be undone."
        ):
            self.queue.clear_all()
            self._refresh_display()

    # =========================================================================
    # Network Folder Scanning
    # =========================================================================

    def _start_network_scan(self):
        """Start background scan of the network manifest folder."""
        app_config = get_config()
        scan_dir = app_config.default_output_dir

        if not scan_dir:
            return

        thread = threading.Thread(
            target=self._network_scan_thread,
            args=(scan_dir,),
            daemon=True
        )
        thread.start()

    def _network_scan_thread(self, scan_dir: str):
        """Background thread: scan network folder for existing manifests."""
        if not is_network_path_accessible(scan_dir):
            self.parent.after(0, self._on_network_scan_complete, 0, 0, False)
            return

        found = scan_manifests(scan_dir)
        added = 0
        skipped = 0

        for m in found:
            result = self.queue.add_manifest_if_new(
                carrier=m["carrier"],
                po_number=m["po_number"],
                manifest_path=m["path"],
            )
            if result is not None:
                added += 1
            else:
                skipped += 1

        self.parent.after(0, self._on_network_scan_complete, added, skipped, True)

    def _on_network_scan_complete(self, added: int, skipped: int, network_ok: bool):
        """Handle network scan completion on the main thread."""
        if not network_ok:
            messagebox.showwarning(
                "Network Folder Unreachable",
                "Could not access the manifest network folder.\n\n"
                "Please check your VPN connection and try refreshing."
            )
            self._log("[!] Network scan: folder unreachable (check VPN)")
            return

        if added > 0:
            self._log(f"Network scan: added {added} manifest(s), {skipped} already in queue")
            self._refresh_display()
        elif skipped > 0:
            self._log(f"Network scan: all {skipped} manifest(s) already in queue")


class GlobalSettingsDialog:
    """Dialog for global pre-alert settings."""

    def __init__(self, parent, config: PreAlertConfig):
        self.result = None
        self.config = config

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Pre-Alert Settings")
        self.dialog.geometry("450x320")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (450 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (320 // 2)
        self.dialog.geometry(f"+{x}+{y}")

        self._create_widgets()

    def _create_widgets(self):
        """Build the dialog UI."""
        main_frame = ttk.Frame(self.dialog, padding="15")
        main_frame.pack(fill='both', expand=True)

        # Sender name
        ttk.Label(main_frame, text="Sender name (for email signature):").pack(anchor='w')
        self.sender_var = tk.StringVar(value=self.config.sender_name)
        ttk.Entry(main_frame, textvariable=self.sender_var, width=50).pack(fill='x', pady=(0, 15))

        # Template path
        ttk.Label(main_frame, text="Email template path:").pack(anchor='w')
        self.template_var = tk.StringVar(value=self.config.email_template_path)
        ttk.Entry(main_frame, textvariable=self.template_var, width=50).pack(fill='x', pady=(0, 5))

        ttk.Label(
            main_frame,
            text="Relative to application directory",
            font=('Helvetica', 8),
            foreground='gray'
        ).pack(anchor='w', pady=(0, 15))

        # Queue settings section
        queue_frame = ttk.LabelFrame(main_frame, text="Queue Settings", padding="10")
        queue_frame.pack(fill='x', pady=(0, 15))

        # Retention days
        retention_frame = ttk.Frame(queue_frame)
        retention_frame.pack(fill='x', pady=(0, 5))

        ttk.Label(retention_frame, text="Retention days:").pack(side='left')
        self.retention_var = tk.IntVar(value=self.config.queue_settings.retention_days)
        ttk.Spinbox(
            retention_frame,
            from_=1,
            to=90,
            textvariable=self.retention_var,
            width=5
        ).pack(side='left', padx=(5, 0))

        # Auto cleanup checkbox
        self.auto_cleanup_var = tk.BooleanVar(value=self.config.queue_settings.auto_cleanup)
        ttk.Checkbutton(
            queue_frame,
            text="Auto-cleanup old entries on startup",
            variable=self.auto_cleanup_var
        ).pack(anchor='w', pady=(5, 0))

        # Expand today checkbox
        self.expand_today_var = tk.BooleanVar(value=self.config.queue_settings.expand_today_on_start)
        ttk.Checkbutton(
            queue_frame,
            text="Expand today's manifests on startup",
            variable=self.expand_today_var
        ).pack(anchor='w')

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x')

        ttk.Button(button_frame, text="Save", command=self._save).pack(side='right', padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self._cancel).pack(side='right')

    def _save(self):
        """Save and close."""
        self.config.sender_name = self.sender_var.get().strip() or "Citipost International Operations"
        self.config.email_template_path = self.template_var.get().strip() or "templates/pre_alert_email.html"

        self.config.queue_settings = QueueSettings(
            retention_days=self.retention_var.get(),
            auto_cleanup=self.auto_cleanup_var.get(),
            expand_today_on_start=self.expand_today_var.get(),
        )

        save_pre_alert_config(self.config)
        self.result = self.config
        self.dialog.destroy()

    def _cancel(self):
        """Close without saving."""
        self.result = None
        self.dialog.destroy()

    def show(self) -> Optional[PreAlertConfig]:
        """Show dialog and return result."""
        self.dialog.wait_window()
        return self.result
