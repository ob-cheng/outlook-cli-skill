"""Email data model."""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass
class Email:
    """Represents an email message."""

    subject: str
    sender: str
    sender_clean: str
    sender_domain: Optional[str]
    sender_smtp: Optional[str]
    to: str
    to_names: list[str]
    to_emails: list[str]
    cc: Optional[str]
    cc_emails: list[str]
    date: Optional[datetime]
    html_body: Optional[str]
    text_body: Optional[str]
    is_sent: bool = False
    is_read: bool = True
    importance: str = "normal"
    has_attachments: bool = False
    message_id: Optional[str] = None

    @classmethod
    def from_mail_item(cls, mail_item, is_sent: bool = False) -> 'Email':
        """Create Email from Outlook MailItem COM object."""
        from ..utils.formatting import (
            extract_email_address,
            extract_display_name,
            get_sender_smtp_address,
            extract_recipients,
        )

        try:
            subject = mail_item.Subject or "No Subject"
        except Exception:
            subject = "No Subject"

        # Sender info
        try:
            sender_name = mail_item.SenderName or ""
            sender_email = get_sender_smtp_address(mail_item) or ""
            sender_smtp = extract_email_address(sender_email)
            sender_clean = extract_display_name(sender_name) or sender_email
            sender_domain = sender_smtp.split('@')[1] if sender_smtp and '@' in sender_smtp else None
            if sender_name and sender_email:
                sender = f"{sender_name} <{sender_email}>"
            else:
                sender = sender_email or sender_name or "Unknown"
        except Exception:
            sender = "Unknown"
            sender_clean = "Unknown"
            sender_domain = None
            sender_smtp = None

        # Recipients
        try:
            to, to_names, to_emails = extract_recipients(mail_item.Recipients)
        except Exception:
            to = "N/A"
            to_names = []
            to_emails = []

        # CC
        try:
            from ..utils.formatting import get_smtp_address
            cc = mail_item.CC if mail_item.CC else None
            cc_emails = []
            for i in range(1, mail_item.Recipients.Count + 1):
                recip = mail_item.Recipients.Item(i)
                if recip.Type == 2:  # CC recipient
                    smtp = get_smtp_address(recip)
                    email = extract_email_address(smtp)
                    if email:
                        cc_emails.append(email)
        except Exception:
            cc = None
            cc_emails = []

        # Date
        try:
            date = mail_item.SentOn if is_sent else mail_item.ReceivedTime
        except Exception:
            date = None

        # Body
        try:
            html_body = mail_item.HTMLBody
        except Exception:
            html_body = None

        try:
            text_body = mail_item.Body
        except Exception:
            text_body = None

        # Additional properties
        try:
            is_read = mail_item.UnRead == False
        except Exception:
            is_read = True

        try:
            importance = str(mail_item.Importance)
            if importance == "2":
                importance = "high"
            elif importance == "0":
                importance = "low"
            else:
                importance = "normal"
        except Exception:
            importance = "normal"

        try:
            has_attachments = mail_item.Attachments.Count > 0
        except Exception:
            has_attachments = False

        try:
            message_id = mail_item.EntryID
        except Exception:
            message_id = None

        return cls(
            subject=subject,
            sender=sender,
            sender_clean=sender_clean,
            sender_domain=sender_domain,
            sender_smtp=sender_smtp,
            to=to,
            to_names=to_names,
            to_emails=to_emails,
            cc=cc,
            cc_emails=cc_emails,
            date=date,
            html_body=html_body,
            text_body=text_body,
            is_sent=is_sent,
            is_read=is_read,
            importance=importance,
            has_attachments=has_attachments,
            message_id=message_id,
        )

    def to_dict(self, include_body: bool = True) -> dict:
        """Convert to dictionary for JSON serialization."""
        result = {
            'message_id': self.message_id,
            'subject': self.subject,
            'sender': self.sender,
            'sender_clean': self.sender_clean,
            'sender_smtp': self.sender_smtp,
            'to': self.to,
            'to_emails': self.to_emails,
            'cc': self.cc,
            'cc_emails': self.cc_emails,
            'date': self.date.isoformat() if self.date else None,
            'is_sent': self.is_sent,
            'is_read': self.is_read,
            'importance': self.importance,
            'has_attachments': self.has_attachments,
        }
        if include_body:
            result['text_body'] = self.text_body
            result['html_body'] = self.html_body
        return result
