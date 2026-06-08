"""Email search service."""

import re
from datetime import datetime
from typing import Optional

from ..core.models import Email
from ..core.folders import find_folder_by_name, find_folder_by_path
from ..utils.formatting import format_outlook_date


class SearchService:
    """Service for searching and filtering emails from Outlook."""

    def __init__(self, namespace):
        """Initialize with Outlook MAPI namespace."""
        self.namespace = namespace

    def search(
        self,
        folders: list[str] | None = None,
        since_date: datetime | None = None,
        until_date: datetime | None = None,
        unread_only: bool = False,
        filter_emails: list[str] | None = None,
        filter_domains: list[str] | None = None,
        filter_keyword: str | None = None,
        max_recipients: int = 20,
    ) -> list[Email]:
        """Search emails across folders with filters.

        Args:
            folders: List of folder names to search. Defaults to Inbox + Sent Items.
            since_date: Only include emails after this date.
            until_date: Only include emails before this date.
            unread_only: If True, only include unread emails.
            filter_emails: List of email addresses to filter by (matches From/To/CC).
            filter_domains: List of domains to filter by.
            filter_keyword: Keyword to search in subject/body.
            max_recipients: Skip emails with more recipients (mass distribution).

        Returns:
            List of Email objects matching the criteria.
        """
        all_emails = []

        if folders:
            for folder_name in folders:
                folder = self._find_folder(folder_name)
                if folder:
                    is_sent = folder_name.lower() in ['sent', 'sent items', 'outbox']
                    emails = self._extract_from_folder(
                        folder, since_date, until_date, is_sent,
                        filter_emails, filter_domains, filter_keyword,
                        unread_only, max_recipients
                    )
                    all_emails.extend(emails)
        else:
            # Default: Inbox + Sent Items
            inbox = self.namespace.GetDefaultFolder(6)
            sent = self.namespace.GetDefaultFolder(5)

            all_emails.extend(self._extract_from_folder(
                inbox, since_date, until_date, False,
                filter_emails, filter_domains, filter_keyword,
                unread_only, max_recipients
            ))
            all_emails.extend(self._extract_from_folder(
                sent, since_date, until_date, True,
                filter_emails, filter_domains, filter_keyword,
                unread_only, max_recipients
            ))

        return all_emails

    def get_message_by_id(self, message_id: str) -> Email | None:
        """Get a single email by its EntryID.

        Args:
            message_id: The Outlook EntryID of the message.

        Returns:
            Email object or None if not found.
        """
        try:
            mail_item = self.namespace.GetItemFromID(message_id)
            if mail_item and mail_item.Class == 43:  # olMail
                return Email.from_mail_item(mail_item)
        except Exception:
            pass
        return None

    def _find_folder(self, folder_name: str):
        """Find folder by name or path."""
        if '/' in folder_name:
            return find_folder_by_path(self.namespace, folder_name)
        return find_folder_by_name(self.namespace, folder_name)

    def _extract_from_folder(
        self,
        folder,
        since_date: datetime | None,
        until_date: datetime | None,
        is_sent: bool,
        filter_emails: list[str] | None,
        filter_domains: list[str] | None,
        filter_keyword: str | None,
        unread_only: bool,
        max_recipients: int,
    ) -> list[Email]:
        """Extract emails from a single folder."""
        emails = []

        try:
            items = folder.Items
            date_field = "[SentOn]" if is_sent else "[ReceivedTime]"

            # Build filter
            filters = []
            if since_date:
                filters.append(f"{date_field} > '{format_outlook_date(since_date)}'")
            if until_date:
                filters.append(f"{date_field} < '{format_outlook_date(until_date)}'")
            if unread_only:
                filters.append("[UnRead] = True")

            if filters:
                filter_str = " AND ".join(filters)
                filtered_items = items.Restrict(filter_str)
            else:
                filtered_items = items

            # Sort by date
            filtered_items.Sort(date_field, True)

            for i in range(1, filtered_items.Count + 1):
                try:
                    mail = filtered_items.Item(i)
                    if mail.Class != 43:  # Not a mail item
                        continue

                    email = Email.from_mail_item(mail, is_sent)

                    # Skip auto-replies
                    if self._is_auto_reply(email.subject):
                        continue

                    # Apply participant filter
                    if (filter_emails or filter_domains):
                        if not self._matches_email_filter(email, filter_emails, filter_domains, max_recipients):
                            continue

                    # Apply keyword filter
                    if filter_keyword and not self._matches_keyword(email, filter_keyword):
                        continue

                    emails.append(email)

                except Exception:
                    continue

        except Exception:
            pass

        return emails

    def _is_auto_reply(self, subject: str) -> bool:
        """Check if email is an automatic reply."""
        if not subject:
            return False
        patterns = [
            r'^Automatic reply:',
            r'^Out of Office:',
            r'^Out of Office Re:',
            r'^Auto:',
            r'^Autoreply:',
            r'^Auto-reply:',
            r'^OOO:',
            r'^Absence:',
        ]
        for pattern in patterns:
            if re.match(pattern, subject, re.IGNORECASE):
                return True
        return False

    def _matches_email_filter(
        self,
        email: Email,
        filter_emails: list[str] | None,
        filter_domains: list[str] | None,
        max_recipients: int,
    ) -> bool:
        """Check if email matches email/domain filter."""
        # Skip mass distribution
        total_recipients = len(email.to_emails) + len(email.cc_emails)
        if total_recipients > max_recipients:
            return False

        filter_email_set = {e.lower().strip() for e in filter_emails} if filter_emails else set()
        filter_domain_set = {d.lower().strip().lstrip('@') for d in filter_domains} if filter_domains else set()

        def matches_domain(addr: str) -> bool:
            if not addr or '@' not in addr:
                return False
            domain = addr.lower().split('@')[1]
            return domain in filter_domain_set

        # Check sender
        if email.sender_smtp:
            if email.sender_smtp.lower() in filter_email_set:
                return True
            if matches_domain(email.sender_smtp):
                return True

        # Check To
        for addr in email.to_emails:
            if addr.lower() in filter_email_set:
                return True
            if matches_domain(addr):
                return True

        # Check CC
        for addr in email.cc_emails:
            if addr.lower() in filter_email_set:
                return True
            if matches_domain(addr):
                return True

        return False

    def _matches_keyword(self, email: Email, keyword: str) -> bool:
        """Check if email contains keyword as a whole word in subject or body."""
        pattern = re.compile(r'\b' + re.escape(keyword) + r'\b', re.IGNORECASE)
        if pattern.search(email.subject or ''):
            return True
        if pattern.search(email.text_body or ''):
            return True
        return False
