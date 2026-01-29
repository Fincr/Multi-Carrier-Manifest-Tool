"""
Pre-Alert Tab UI
================
Tkinter tab for managing and sending pre-alert emails.
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
from typing import Dict, List, Optional, Callable
import threading

from .config_manager import PreAlertConfig, CarrierEmailConfig, load_pre_alert_config, save_pre_alert_config
from .send_tracker import SendTracker, SendRecord
from .email_sender import send_pre_alert_email


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
        
        # Manifest data: {carrier_name: {po, path, date}}
        self.manifests: Dict[str, dict] = {}
        
        # Checkbox variables
        self.check_vars: Dict[str, tk.BooleanVar] = {}
        
        # Log callback (set by main app)
        self.log_callback: Optional[Callable[[str], None]] = None
        
        self._create_widgets()
        self._refresh_display()
    
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
            text="⚙ Settings",
            command=self._open_global_settings
        ).pack(side='right')
        
        ttk.Label(
            main_frame,
            text="Select manifests to send pre-alert emails. Only carriers requiring pre-alerts are shown.",
            foreground='gray'
        ).pack(anchor='w', pady=(0, 10))
        
        # Manifest list frame
        list_frame = ttk.LabelFrame(main_frame, text="Available Manifests", padding="10")
        list_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Create treeview for manifest list
        columns = ('carrier', 'po', 'file', 'status')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)
        
        self.tree.heading('carrier', text='Carrier')
        self.tree.heading('po', text='PO Number')
        self.tree.heading('file', text='Manifest File')
        self.tree.heading('status', text='Status')
        
        self.tree.column('carrier', width=120)
        self.tree.column('po', width=80)
        self.tree.column('file', width=250)
        self.tree.column('status', width=100)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Selection buttons
        select_frame = ttk.Frame(main_frame)
        select_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Button(select_frame, text="Select All", command=self._select_all).pack(side='left', padx=(0, 5))
        ttk.Button(select_frame, text="Select None", command=self._select_none).pack(side='left', padx=(0, 5))
        ttk.Button(select_frame, text="Configure Selected", command=self._configure_selected).pack(side='left', padx=(0, 5))
        ttk.Button(select_frame, text="Refresh", command=self._refresh_display).pack(side='right')
        
        # Action buttons
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill='x', pady=(0, 10))
        
        self.send_button = ttk.Button(
            action_frame,
            text="📧 Send Pre-Alerts",
            command=self._send_selected
        )
        self.send_button.pack(side='left', padx=(0, 10))
        
        ttk.Button(
            action_frame,
            text="Clear Sent Status",
            command=self._clear_sent_status
        ).pack(side='left')
        
        # Log area
        log_frame = ttk.LabelFrame(main_frame, text="Send Log", padding="5")
        log_frame.pack(fill='both', expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, font=('Consolas', 9))
        self.log_text.pack(fill='both', expand=True)
    
    def _refresh_display(self):
        """Refresh the manifest list display."""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Re-check sent status
        today_sent = self.tracker.get_all_today()
        
        # Add manifests
        for carrier, data in self.manifests.items():
            canonical = self.config.get_canonical_carrier_name(carrier)
            if not canonical:
                continue
            
            # Check if already sent today
            if canonical in today_sent:
                status = "✓ Sent"
            else:
                carrier_config = self.config.get_carrier_config(carrier)
                if carrier_config and carrier_config.enabled and carrier_config.recipients:
                    status = "Ready"
                elif carrier_config and not carrier_config.enabled:
                    status = "Disabled"
                else:
                    status = "Not configured"
            
            self.tree.insert('', 'end', iid=carrier, values=(
                canonical,
                data.get('po', ''),
                os.path.basename(data.get('path', '')),
                status
            ))
    
    def _select_all(self):
        """Select all items in the tree."""
        for item in self.tree.get_children():
            self.tree.selection_add(item)
    
    def _select_none(self):
        """Deselect all items."""
        self.tree.selection_remove(*self.tree.get_children())
    
    def _configure_selected(self):
        """Open configuration dialog for selected carrier."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a carrier to configure.")
            return
        
        # Get first selected carrier
        carrier = selected[0]
        canonical = self.config.get_canonical_carrier_name(carrier)
        if not canonical:
            return
        
        # Open config dialog
        dialog = PreAlertConfigDialog(self.parent, self.config, canonical)
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
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one manifest to send.")
            return
        
        # Validate selected manifests
        to_send = []
        for carrier in selected:
            canonical = self.config.get_canonical_carrier_name(carrier)
            if not canonical:
                continue
            
            carrier_config = self.config.get_carrier_config(carrier)
            if not carrier_config or not carrier_config.enabled:
                self._log(f"⚠ {canonical}: Skipped (disabled)")
                continue
            
            if not carrier_config.recipients:
                self._log(f"⚠ {canonical}: Skipped (no recipients configured)")
                continue
            
            if canonical not in self.manifests:
                self._log(f"⚠ {canonical}: Skipped (no manifest data)")
                continue
            
            # Check if already sent
            if self.tracker.was_sent_today(canonical):
                if not messagebox.askyesno(
                    "Already Sent",
                    f"Pre-alert for {canonical} was already sent today.\n\nSend again?"
                ):
                    self._log(f"⚠ {canonical}: Skipped (already sent today)")
                    continue
            
            to_send.append((carrier, canonical, carrier_config))
        
        if not to_send:
            messagebox.showinfo("Nothing to Send", "No valid manifests to send.")
            return
        
        # Confirm
        msg = f"Send pre-alerts for {len(to_send)} manifest(s)?\n\n"
        for _, canonical, _ in to_send:
            msg += f"• {canonical}\n"
        
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
        
        for carrier, canonical, carrier_config in to_send:
            manifest_data = self.manifests.get(carrier, {})
            
            self.parent.after(0, self._log, f"\nSending pre-alert for {canonical}...")
            
            # Get template path
            template_path = os.path.join(self.app_dir, self.config.email_template_path)
            
            # Send email
            success, message = send_pre_alert_email(
                carrier_name=canonical,
                po_number=manifest_data.get('po', ''),
                manifest_path=manifest_data.get('path', ''),
                recipients=carrier_config.recipients,
                cc=carrier_config.cc,
                subject_template=carrier_config.subject_template,
                sender_name=self.config.sender_name,
                template_path=template_path
            )
            
            # Record result
            record = SendRecord(
                po_number=manifest_data.get('po', ''),
                manifest_path=manifest_data.get('path', ''),
                sent_at=datetime.now().strftime("%H:%M:%S"),
                recipients=carrier_config.recipients,
                cc=carrier_config.cc,
                success=success,
                error_message="" if success else message
            )
            self.tracker.record_send(canonical, record)
            
            results.append((canonical, success, message))
            
            if success:
                self.parent.after(0, self._log, f"  ✓ {message}")
                self.parent.after(0, self._log, f"    To: {', '.join(carrier_config.recipients)}")
                if carrier_config.cc:
                    self.parent.after(0, self._log, f"    CC: {', '.join(carrier_config.cc)}")
            else:
                self.parent.after(0, self._log, f"  ✗ {message}")
        
        # Complete
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
    
    def _clear_sent_status(self):
        """Clear today's sent status."""
        if not messagebox.askyesno(
            "Confirm Clear",
            "Clear today's sent status?\n\nThis allows you to re-send pre-alerts."
        ):
            return
        
        self.tracker.clear_today()
        self._refresh_display()
        self._log("Cleared today's sent status")
    
    def _log(self, message: str):
        """Add message to the log display."""
        self.log_text.insert('end', message + '\n')
        self.log_text.see('end')
        
        # Also send to main app log if callback set
        if self.log_callback:
            self.log_callback(message)
    
    def add_manifest(self, carrier_name: str, po_number: str, manifest_path: str):
        """
        Add a processed manifest to the pre-alert list.
        
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
        
        self.manifests[carrier_name] = {
            'po': po_number,
            'path': manifest_path,
            'date': datetime.now().strftime("%Y-%m-%d"),
        }
        
        self._refresh_display()
    
    def clear_manifests(self):
        """Clear all manifests from the list."""
        self.manifests.clear()
        self._refresh_display()


class GlobalSettingsDialog:
    """Dialog for global pre-alert settings."""
    
    def __init__(self, parent, config: PreAlertConfig):
        self.result = None
        self.config = config
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Pre-Alert Settings")
        self.dialog.geometry("450x200")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (450 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (200 // 2)
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
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x')
        
        ttk.Button(button_frame, text="Save", command=self._save).pack(side='right', padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self._cancel).pack(side='right')
    
    def _save(self):
        """Save and close."""
        self.config.sender_name = self.sender_var.get().strip() or "Citipost International Operations"
        self.config.email_template_path = self.template_var.get().strip() or "templates/pre_alert_email.html"
        
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
