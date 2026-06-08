"""Formatting and parsing utilities."""

import re
from datetime import datetime


def extract_email_address(full_address: str) -> str | None:
    """Extract just the email address from a full address string."""
    if not full_address:
        return None
    match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', str(full_address))
    return match.group(0).lower() if match else None


def extract_all_email_addresses(field_string: str) -> list[str]:
    """Extract all email addresses from a To/CC field string."""
    if not field_string:
        return []
    matches = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', str(field_string))
    return [m.lower() for m in matches]


def extract_display_name(full_address: str) -> str | None:
    """Extract display name from address, cleaning up Exchange paths."""
    if not full_address:
        return None
    name = str(full_address)

    if '/CN=' in name.upper():
        parts = name.split('/')
        for part in reversed(parts):
            if part.upper().startswith('CN='):
                name = part[3:]
                break

    name = re.sub(r'\s*<[^>]+>\s*', '', name)
    name = re.sub(r'\([^)]*\)', '', name)
    name = name.strip()

    if ',' in name:
        parts = [p.strip() for p in name.split(',')]
        if len(parts) == 2 and parts[0] and parts[1]:
            if (parts[0][0].isupper() and parts[1][0].isupper() and
                len(parts[0]) < 20 and len(parts[1]) < 20):
                name = f"{parts[1]} {parts[0]}"

    return name if name else None


def get_smtp_address(recip) -> str:
    """Get SMTP email address from recipient, handling Exchange format."""
    try:
        addr_entry = recip.AddressEntry
        if addr_entry.Type == "EX":
            exch_user = addr_entry.GetExchangeUser()
            if exch_user:
                return exch_user.PrimarySmtpAddress
        return recip.Address
    except Exception:
        return recip.Address


def get_sender_smtp_address(mail_item) -> str:
    """Get SMTP email address of sender, handling Exchange format."""
    try:
        sender = mail_item.Sender
        if sender and sender.Type == "EX":
            exch_user = sender.GetExchangeUser()
            if exch_user:
                return exch_user.PrimarySmtpAddress
        return mail_item.SenderEmailAddress
    except Exception:
        return mail_item.SenderEmailAddress


def extract_recipients(recipients_obj) -> tuple[str, list[str], list[str]]:
    """Extract recipient information from Outlook Recipients object.

    Returns:
        tuple: (formatted_string, list_of_names, list_of_emails)
    """
    if not recipients_obj:
        return "N/A", [], []
    try:
        recipients = []
        names = []
        emails = []
        for i in range(1, recipients_obj.Count + 1):
            recip = recipients_obj.Item(i)
            name = recip.Name
            email = get_smtp_address(recip)
            if name:
                clean_name = extract_display_name(name)
                if clean_name:
                    names.append(clean_name)
            if email:
                smtp_email = extract_email_address(email)
                if smtp_email:
                    emails.append(smtp_email)
            if name and email:
                recipients.append(f"{name} <{email}>")
            elif email:
                recipients.append(email)
            elif name:
                recipients.append(name)
        return "; ".join(recipients) if recipients else "N/A", names, emails
    except Exception:
        return "N/A", [], []


def format_outlook_date(dt: datetime) -> str:
    """Format datetime for Outlook Restrict filter."""
    return dt.strftime("%m/%d/%Y %H:%M %p")


def sanitize_filename(filename: str, max_length: int = 50) -> str:
    """Sanitize filename by removing invalid characters and limiting length."""
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    filename = re.sub(r'\s+', ' ', filename)
    filename = filename.strip('. ')
    if not filename:
        return "untitled"
    if len(filename) > max_length:
        filename = filename[:max_length].strip()
    return filename


def normalize_subject(subject: str) -> str:
    """Normalize subject by removing RE:/FW: prefixes for thread grouping."""
    if not subject:
        return "No Subject"
    normalized = re.sub(r'^(RE|FW|Fwd|Re|Fw):\s*', '', subject, flags=re.IGNORECASE)
    normalized = re.sub(r'^(RE|FW|Fwd|Re|Fw):\s*', '', normalized, flags=re.IGNORECASE)
    return normalized.strip() or "No Subject"


def parse_date(date_str: str) -> datetime:
    """Parse a date string in various formats."""
    formats = ['%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%d-%m-%Y']
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Cannot parse date: {date_str}. Use YYYY-MM-DD format.")
