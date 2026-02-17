"""
Outlook Email Sender
====================
Sends pre-alert emails via Outlook COM automation.
"""

import os
import re
from datetime import datetime
from typing import List, Optional, Tuple
from dataclasses import dataclass


@dataclass
class EmailResult:
    """Result of an email send attempt."""
    success: bool
    message: str
    recipients: List[str]
    cc: List[str]


class OutlookEmailSender:
    """
    Sends emails via Outlook desktop client using COM automation.
    
    Emails are sent from the user's default Outlook account,
    appear in their Sent folder, and use their signature if configured.
    """
    
    def __init__(self):
        self.outlook = None
        self._initialized = False
    
    def _ensure_outlook(self) -> bool:
        """
        Initialize Outlook COM connection.
        
        Returns:
            True if Outlook is available, False otherwise.
        """
        if self._initialized:
            return self.outlook is not None
        
        self._initialized = True
        
        try:
            import win32com.client
            import pythoncom
            
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            # Get Outlook application
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            return True
            
        except ImportError:
            print("Warning: pywin32 not installed. Run: pip install pywin32")
            return False
        except Exception as e:
            print(f"Warning: Could not connect to Outlook: {e}")
            return False
    
    def send_email(
        self,
        recipients: List[str],
        cc: List[str],
        subject: str,
        body_html: str,
        attachments: List[str],
        display_only: bool = False
    ) -> EmailResult:
        """
        Send an email via Outlook.
        
        Args:
            recipients: List of TO email addresses
            cc: List of CC email addresses
            subject: Email subject line
            body_html: HTML body content
            attachments: List of file paths to attach
            display_only: If True, display email for review instead of sending
        
        Returns:
            EmailResult with success status and details
        """
        if not self._ensure_outlook():
            return EmailResult(
                success=False,
                message="Outlook not available. Please ensure Outlook is installed and running.",
                recipients=recipients,
                cc=cc
            )
        
        if not recipients:
            return EmailResult(
                success=False,
                message="No recipients specified",
                recipients=recipients,
                cc=cc
            )
        
        try:
            # Create mail item (0 = olMailItem)
            mail = self.outlook.CreateItem(0)
            
            # Set recipients
            mail.To = "; ".join(recipients)
            if cc:
                mail.CC = "; ".join(cc)
            
            # Set subject and body
            mail.Subject = subject
            mail.HTMLBody = body_html
            
            # Add attachments
            for filepath in attachments:
                if os.path.exists(filepath):
                    mail.Attachments.Add(filepath)
                else:
                    return EmailResult(
                        success=False,
                        message=f"Attachment not found: {filepath}",
                        recipients=recipients,
                        cc=cc
                    )
            
            # Send or display
            if display_only:
                mail.Display()
                return EmailResult(
                    success=True,
                    message="Email displayed for review",
                    recipients=recipients,
                    cc=cc
                )
            else:
                mail.Send()
                return EmailResult(
                    success=True,
                    message=f"Email sent to {len(recipients)} recipient(s)",
                    recipients=recipients,
                    cc=cc
                )
        
        except Exception as e:
            return EmailResult(
                success=False,
                message=f"Failed to send email: {str(e)}",
                recipients=recipients,
                cc=cc
            )
    
    def cleanup(self):
        """Release COM resources."""
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass
        
        self.outlook = None
        self._initialized = False


def load_email_template(template_path: str) -> str:
    """
    Load HTML email template from file.
    
    Args:
        template_path: Path to the HTML template file
    
    Returns:
        Template content as string, or default template if not found.
    """
    if os.path.exists(template_path):
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                return f.read()
        except IOError:
            pass
    
    # Return default template if file not found
    return get_default_template()


def get_default_template() -> str:
    """Return the default HTML email template."""
    return """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body {
            font-family: Arial, sans-serif;
            font-size: 14px;
            line-height: 1.5;
            color: #333;
        }
        .header {
            margin-bottom: 20px;
        }
        .content {
            margin-bottom: 20px;
        }
        .details {
            background-color: #f5f5f5;
            padding: 15px;
            border-radius: 5px;
            margin: 15px 0;
        }
        .details table {
            width: 100%;
            border-collapse: collapse;
        }
        .details td {
            padding: 5px 10px;
        }
        .details td:first-child {
            font-weight: bold;
            width: 120px;
        }
        .footer {
            margin-top: 30px;
            color: #666;
        }
    </style>
</head>
<body>
    <div class="header">
        <p>Dear Team,</p>
    </div>
    
    <div class="content">
        <p>Please find attached today's manifest for <strong>{carrier}</strong>.</p>
        
        <div class="details">
            <table>
                <tr>
                    <td>Dispatch Date:</td>
                    <td>{date}</td>
                </tr>
                <tr>
                    <td>PO Number:</td>
                    <td>{po_number}</td>
                </tr>
            </table>
        </div>
    </div>
    
    <div class="footer">
        <p>Kind regards,<br>
        {sender_name}</p>
    </div>
</body>
</html>"""


def format_email_body(
    template: str,
    carrier: str,
    date: str,
    po_number: str,
    sender_name: str
) -> str:
    """
    Format the email template with provided values.
    
    Args:
        template: HTML template with placeholders
        carrier: Carrier name
        date: Dispatch date
        po_number: PO number
        sender_name: Sender name for signature
    
    Returns:
        Formatted HTML content
    """
    # Use replace() instead of format() to avoid conflicts with CSS curly braces
    return (template
        .replace("{carrier}", carrier)
        .replace("{date}", date)
        .replace("{po_number}", po_number)
        .replace("{sender_name}", sender_name)
    )


def find_companion_files(manifest_path: str) -> List[str]:
    """
    Find companion PDF files in the same directory as the manifest.

    Matches PDFs that share the same carrier prefix and PO number
    as the Excel manifest filename.

    Args:
        manifest_path: Path to the Excel manifest file.

    Returns:
        List of paths to matching PDF files (empty if none found).
    """
    directory = os.path.dirname(manifest_path)
    filename = os.path.basename(manifest_path)

    # Extract carrier prefix (leading non-numeric underscore-separated words)
    # e.g. "Deutsche_Post" from "Deutsche_Post_5367_27911_16.02.2026_20260216_131915.xlsx"
    prefix_match = re.match(r'^([A-Za-z][A-Za-z_-]*?)_\d', filename)
    if not prefix_match:
        return []
    carrier_prefix = prefix_match.group(1)

    # Extract all 5-digit numbers as candidate PO numbers
    po_numbers = re.findall(r'_(\d{5})(?:_|\.)', filename)
    if not po_numbers:
        return []

    companions = []
    try:
        for f in os.listdir(directory):
            if not f.lower().endswith('.pdf'):
                continue
            if not f.startswith(carrier_prefix + '_'):
                continue
            # Check that at least one PO number appears in the PDF filename
            if any(f'_{po}_' in f or f'_{po}.' in f for po in po_numbers):
                companions.append(os.path.join(directory, f))
    except OSError:
        pass

    return companions


def send_pre_alert_email(
    carrier_name: str,
    po_number: str,
    manifest_path: str,
    recipients: List[str],
    cc: List[str],
    subject_template: str,
    sender_name: str,
    template_path: Optional[str] = None,
    display_only: bool = False
) -> Tuple[bool, str]:
    """
    Send a pre-alert email for a carrier manifest.
    
    This is the main entry point for sending pre-alert emails.
    
    Args:
        carrier_name: Name of the carrier
        po_number: PO number for the manifest
        manifest_path: Path to the manifest file to attach
        recipients: List of TO email addresses
        cc: List of CC email addresses
        subject_template: Subject line template (can include {carrier}, {date}, {po_number})
        sender_name: Name to use in signature
        template_path: Path to HTML template file (optional)
        display_only: If True, display email instead of sending
    
    Returns:
        (success, message) tuple
    """
    # Get today's date
    today = datetime.now().strftime("%d/%m/%Y")
    
    # Format subject
    subject = subject_template.format(
        carrier=carrier_name,
        date=today,
        po_number=po_number
    )
    
    # Load and format body
    template = load_email_template(template_path) if template_path else get_default_template()
    body_html = format_email_body(
        template=template,
        carrier=carrier_name,
        date=today,
        po_number=po_number,
        sender_name=sender_name
    )
    
    # Send email
    sender = OutlookEmailSender()
    try:
        result = sender.send_email(
            recipients=recipients,
            cc=cc,
            subject=subject,
            body_html=body_html,
            attachments=[manifest_path] + find_companion_files(manifest_path),
            display_only=display_only
        )
        return result.success, result.message
    finally:
        sender.cleanup()
