"""Markdown export service."""

import re
import json
import hashlib
from pathlib import Path
from datetime import datetime
from collections import defaultdict

from bs4 import BeautifulSoup
from markdownify import markdownify as md

from ..core.models import Email
from ..utils.formatting import sanitize_filename, normalize_subject


STATE_FILE = "extraction_state.json"


def clean_surrogates(text: str) -> str:
    """Remove surrogate characters that can't be encoded to UTF-8."""
    if not text:
        return text
    return text.encode('utf-8', errors='replace').decode('utf-8')


class ExportService:
    """Service for exporting emails to Obsidian-flavored markdown."""

    def __init__(self, output_dir: str | Path, state_dir: str | Path | None = None):
        """Initialize with output directory.

        Args:
            output_dir: Directory to save markdown files.
            state_dir: Directory for state file (defaults to script directory).
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        # Store state in a fixed location, not in the output directory
        if state_dir:
            self.state_file = Path(state_dir) / STATE_FILE
        else:
            # Default to the directory where outlook.py lives
            self.state_file = Path(__file__).parent.parent.parent / STATE_FILE

    def get_last_run(self) -> datetime | None:
        """Get the timestamp of the last successful export."""
        if not self.state_file.exists():
            return None
        try:
            data = json.loads(self.state_file.read_text(encoding='utf-8'))
            # Check both keys for backwards compatibility
            last_run = data.get('last_run') or data.get('last_extraction')
            if last_run:
                return datetime.fromisoformat(last_run)
        except Exception:
            pass
        return None

    def save_state(self) -> None:
        """Save the current timestamp as last run."""
        data = {'last_run': datetime.now().isoformat()}
        self.state_file.write_text(json.dumps(data, indent=2), encoding='utf-8')

    def export_emails(
        self,
        emails: list[Email],
        group_threads: bool = True,
        no_overwrite: bool = False,
        save_state: bool = False,
        output_format: str = 'markdown',
        batch: bool = False,
    ) -> dict:
        """Export emails to files.

        Args:
            emails: List of Email objects to export.
            group_threads: If True, group emails by subject into threads.
            no_overwrite: If True, skip existing files.
            save_state: If True, save timestamp for incremental exports.
            output_format: 'markdown' or 'json'.
            batch: For JSON, combine all emails into a single file.

        Returns:
            dict with keys: files_created, files_skipped, emails_processed
        """
        result = {
            'files_created': 0,
            'files_skipped': 0,
            'emails_processed': len(emails),
            'files': [],
        }

        if not emails:
            if save_state:
                self.save_state()
            return result

        if output_format == 'json':
            return self._export_json(emails, group_threads, no_overwrite, save_state, batch)

        if group_threads:
            threads = self._group_into_threads(emails)
            for thread_subject, thread_emails in threads.items():
                file_path = self._export_thread(thread_emails, thread_subject, no_overwrite)
                if file_path:
                    result['files_created'] += 1
                    result['files'].append(str(file_path))
                else:
                    result['files_skipped'] += 1
        else:
            for email in emails:
                file_path = self._export_single(email, no_overwrite)
                if file_path:
                    result['files_created'] += 1
                    result['files'].append(str(file_path))
                else:
                    result['files_skipped'] += 1

        if save_state:
            self.save_state()

        return result

    def _export_json(
        self,
        emails: list[Email],
        group_threads: bool,
        no_overwrite: bool,
        save_state: bool,
        batch: bool,
    ) -> dict:
        """Export emails to JSON format."""
        result = {
            'files_created': 0,
            'files_skipped': 0,
            'emails_processed': len(emails),
            'files': [],
        }

        if batch:
            file_path = self._export_json_batch(emails, group_threads, no_overwrite)
            if file_path:
                result['files_created'] = 1
                result['files'].append(str(file_path))
            else:
                result['files_skipped'] = 1
        elif group_threads:
            threads = self._group_into_threads(emails)
            for thread_subject, thread_emails in threads.items():
                file_path = self._export_json_thread(thread_emails, thread_subject, no_overwrite)
                if file_path:
                    result['files_created'] += 1
                    result['files'].append(str(file_path))
                else:
                    result['files_skipped'] += 1
        else:
            for email in emails:
                file_path = self._export_json_single(email, no_overwrite)
                if file_path:
                    result['files_created'] += 1
                    result['files'].append(str(file_path))
                else:
                    result['files_skipped'] += 1

        if save_state:
            self.save_state()

        return result

    def _export_json_batch(
        self,
        emails: list[Email],
        group_threads: bool,
        no_overwrite: bool,
    ) -> Path | None:
        """Export all emails to a single JSON file."""
        filename = f"emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        output_path = self.output_dir / filename

        if output_path.exists() and no_overwrite:
            return None

        if group_threads:
            threads = self._group_into_threads(emails)
            data = {
                "export_date": datetime.now().isoformat(),
                "thread_count": len(threads),
                "email_count": len(emails),
                "threads": [
                    self._thread_to_json(thread_emails, thread_subject)
                    for thread_subject, thread_emails in threads.items()
                ]
            }
        else:
            data = {
                "export_date": datetime.now().isoformat(),
                "email_count": len(emails),
                "emails": [self._email_to_json(e) for e in emails]
            }

        output_path.write_text(clean_surrogates(json.dumps(data, indent=2, ensure_ascii=False)), encoding='utf-8')
        return output_path

    def _export_json_thread(
        self,
        emails: list[Email],
        thread_subject: str,
        no_overwrite: bool,
    ) -> Path | None:
        """Export a thread to a JSON file."""
        sorted_emails = sorted(
            [e for e in emails if e.date],
            key=lambda x: x.date
        )
        if not sorted_emails:
            return None

        latest_date = max(e.date for e in sorted_emails if e.date)
        date_prefix = latest_date.strftime("%Y-%m-%d_%H%M")
        filename = f"{date_prefix}_{sanitize_filename(thread_subject)}.json"
        output_path = self.output_dir / filename

        if output_path.exists():
            if no_overwrite:
                return None
            file_hash = hashlib.md5(thread_subject.encode('utf-8', errors='replace')).hexdigest()[:8]
            filename = f"{date_prefix}_{sanitize_filename(thread_subject)}_{file_hash}.json"
            output_path = self.output_dir / filename
            if output_path.exists() and no_overwrite:
                return None

        data = self._thread_to_json(sorted_emails, thread_subject)
        output_path.write_text(clean_surrogates(json.dumps(data, indent=2, ensure_ascii=False)), encoding='utf-8')
        return output_path

    def _export_json_single(self, email: Email, no_overwrite: bool) -> Path | None:
        """Export a single email to a JSON file."""
        if not email.date:
            return None

        date_prefix = email.date.strftime("%Y-%m-%d_%H%M")
        filename = f"{date_prefix}_{sanitize_filename(email.subject)}.json"
        output_path = self.output_dir / filename

        if output_path.exists() and no_overwrite:
            return None

        data = self._email_to_json(email)
        output_path.write_text(clean_surrogates(json.dumps(data, indent=2, ensure_ascii=False)), encoding='utf-8')
        return output_path

    def _thread_to_json(self, emails: list[Email], thread_subject: str) -> dict:
        """Convert a thread to JSON-friendly dict."""
        sorted_emails = sorted(
            [e for e in emails if e.date],
            key=lambda x: x.date
        )
        if not sorted_emails:
            return {}

        first_date = sorted_emails[0].date
        last_date = sorted_emails[-1].date
        participants = self._collect_participants(sorted_emails)

        return {
            "subject": thread_subject,
            "message_count": len(sorted_emails),
            "date_start": first_date.isoformat() if first_date else None,
            "date_end": last_date.isoformat() if last_date else None,
            "participants": participants,
            "messages": [self._email_to_json(e) for e in sorted_emails]
        }

    def _email_to_json(self, email: Email) -> dict:
        """Convert email to lean JSON dict optimized for AI ingestion."""
        content = self._extract_content(email)
        content = self._clean_content(content)

        return {
            "id": email.message_id,
            "subject": email.subject,
            "from": email.sender_clean,
            "from_email": email.sender_smtp,
            "to": email.to_emails,
            "cc": email.cc_emails if email.cc_emails else None,
            "date": email.date.isoformat() if email.date else None,
            "direction": "sent" if email.is_sent else "received",
            "body": content if content else None,
        }

    def to_json_data(self, emails: list[Email], group_threads: bool = True) -> dict:
        """Generate JSON data structure without writing files.

        Args:
            emails: List of Email objects.
            group_threads: If True, group emails by subject into threads.

        Returns:
            dict ready for JSON serialization.
        """
        if not emails:
            return {
                "export_date": datetime.now().isoformat(),
                "email_count": 0,
                "threads" if group_threads else "emails": []
            }

        if group_threads:
            threads = self._group_into_threads(emails)
            return {
                "export_date": datetime.now().isoformat(),
                "thread_count": len(threads),
                "email_count": len(emails),
                "threads": [
                    self._thread_to_json(thread_emails, thread_subject)
                    for thread_subject, thread_emails in threads.items()
                ]
            }
        else:
            return {
                "export_date": datetime.now().isoformat(),
                "email_count": len(emails),
                "emails": [self._email_to_json(e) for e in emails]
            }

    def _group_into_threads(self, emails: list[Email]) -> dict[str, list[Email]]:
        """Group emails by normalized subject."""
        threads = defaultdict(list)
        for email in emails:
            thread_key = normalize_subject(email.subject)
            threads[thread_key].append(email)
        return dict(threads)

    def _export_thread(
        self,
        emails: list[Email],
        thread_subject: str,
        no_overwrite: bool,
    ) -> Path | None:
        """Export a thread of emails to a single markdown file."""
        # Sort by date
        sorted_emails = sorted(
            [e for e in emails if e.date],
            key=lambda x: x.date
        )
        if not sorted_emails:
            return None

        # Generate filename
        latest_date = max(e.date for e in sorted_emails if e.date)
        date_prefix = latest_date.strftime("%Y-%m-%d %H%M")
        filename = f"{date_prefix} - {sanitize_filename(thread_subject)}.md"
        output_path = self.output_dir / filename

        # Handle duplicates or skip
        if output_path.exists():
            if no_overwrite:
                return None
            file_hash = hashlib.md5(thread_subject.encode('utf-8', errors='replace')).hexdigest()[:8]
            filename = f"{date_prefix} - {sanitize_filename(thread_subject)}_{file_hash}.md"
            output_path = self.output_dir / filename
            if output_path.exists() and no_overwrite:
                return None

        # Generate markdown
        content = self._create_thread_markdown(sorted_emails, thread_subject)
        if not content:
            return None

        output_path.write_text(clean_surrogates(content), encoding='utf-8')
        return output_path

    def _export_single(self, email: Email, no_overwrite: bool) -> Path | None:
        """Export a single email to a markdown file."""
        if not email.date:
            return None

        date_prefix = email.date.strftime("%Y-%m-%d %H%M")
        filename = f"{date_prefix} - {sanitize_filename(email.subject)}.md"
        output_path = self.output_dir / filename

        if output_path.exists() and no_overwrite:
            return None

        content = self._create_single_markdown(email)
        output_path.write_text(clean_surrogates(content), encoding='utf-8')
        return output_path

    def _create_thread_markdown(self, emails: list[Email], thread_subject: str) -> str:
        """Create markdown content for an email thread."""
        first_date = emails[0].date
        last_date = emails[-1].date
        participants = self._collect_participants(emails)
        tags = self._determine_tags(emails)

        # YAML frontmatter
        safe_title = thread_subject.replace('"', "'")
        lines = ["---"]
        lines.append(f'title: "{safe_title}"')
        if first_date:
            lines.append(f"date: {first_date.strftime('%Y-%m-%d')}")
        if first_date and last_date and first_date.date() != last_date.date():
            lines.append(f"date_end: {last_date.strftime('%Y-%m-%d')}")
        lines.append(f"message_count: {len(emails)}")

        if participants:
            lines.append("participants:")
            for p in participants[:10]:
                lines.append(f'  - "{p.replace(chr(34), chr(39))}"')

        lines.append("tags:")
        for tag in tags:
            lines.append(f"  - {tag}")
        lines.append("---\n")

        # Thread summary callout
        date_range = self._format_date_wikilink(first_date)
        if first_date and last_date and first_date.date() != last_date.date():
            date_range = f"{self._format_date_wikilink(first_date)} to {self._format_date_wikilink(last_date)}"

        lines.append("> [!info] Thread Summary")
        lines.append(f"> **{len(emails)} message(s)** from {date_range}")
        participants_str = ', '.join(participants[:5])
        if len(participants) > 5:
            participants_str += f" +{len(participants) - 5} more"
        lines.append(f"> **Participants:** {participants_str}\n")

        # Messages
        for i, email in enumerate(emails, 1):
            lines.extend(self._format_message(email, i))

        # AI hints
        lines.append("%%")
        lines.append("ai-hints:")
        lines.append("  thread_type: email_conversation")
        lines.append(f"  message_count: {len(emails)}")
        lines.append(f"  direction: {'mixed' if 'email/conversation' in tags else ('outbound' if 'email/sent' in tags else 'inbound')}")
        lines.append("%%")

        return "\n".join(lines)

    def _create_single_markdown(self, email: Email) -> str:
        """Create markdown content for a single email."""
        lines = ["---"]
        safe_title = email.subject.replace('"', "'")
        lines.append(f'title: "{safe_title}"')
        if email.date:
            lines.append(f"date: {email.date.strftime('%Y-%m-%d')}")
        lines.append(f"from: \"{email.sender_clean}\"")

        tags = ["email/single"]
        tags.append("email/sent" if email.is_sent else "email/received")
        lines.append("tags:")
        for tag in tags:
            lines.append(f"  - {tag}")
        lines.append("---\n")

        lines.extend(self._format_message(email, 1))
        return "\n".join(lines)

    def _format_message(self, email: Email, index: int) -> list[str]:
        """Format a single message within a thread."""
        lines = []
        direction = "SENT" if email.is_sent else "RECEIVED"
        direction_tag = "#email/sent" if email.is_sent else "#email/received"

        lines.append(f"## Message {index} ({direction}) ^msg-{index}\n")
        lines.append(f"**From:** {email.sender_clean}  ")

        date_link = self._format_date_wikilink(email.date)
        time_str = email.date.strftime('%H:%M') if email.date else ''
        lines.append(f"**Date:** {date_link} {time_str}  ")
        lines.append(f"{direction_tag}\n")

        content = self._extract_content(email)
        content = self._clean_content(content)
        lines.append(content if content else "*No content*")
        lines.append("\n---\n")

        return lines

    def _extract_content(self, email: Email) -> str:
        """Extract and clean content from email."""
        if email.html_body:
            soup = BeautifulSoup(email.html_body, 'html.parser')

            for tag in soup(['script', 'style', 'meta', 'head']):
                tag.decompose()
            for blockquote in soup.find_all('blockquote'):
                blockquote.decompose()

            content = md(str(soup), heading_style="ATX")

            # Strip quoted content
            patterns = [
                r'\n---\s*\n\*\*From:\*\*.*?\*\*Subject:\*\*.*?(?=\n|$).*',
                r'\n\*\*From:\*\*.*?\*\*Sent:\*\*.*?\*\*Subject:\*\*.*',
                r'\nFrom:.*?\nSent:.*?\nSubject:.*',
                r'\n-{3,}\s*\nFrom:.*',
                r'\nOn .+wrote:.*',
            ]
            for pattern in patterns:
                match = re.search(pattern, content, re.IGNORECASE | re.DOTALL)
                if match:
                    content = content[:match.start()]
                    break

        elif email.text_body:
            lines = email.text_body.split('\n')
            new_lines = []
            for line in lines:
                if line.strip().startswith('>'):
                    break
                if re.match(r'^On .+ wrote:$', line.strip()):
                    break
                if re.match(r'^From:', line.strip()):
                    break
                new_lines.append(line)
            content = '\n'.join(new_lines)
        else:
            content = "*No content available*"

        return content

    def _clean_content(self, content: str) -> str:
        """Clean up content for markdown output."""
        # Remove images
        content = re.sub(r'\[!\[[^\]]*\]\([^\)]*\)\]\([^\)]*\)', '', content)
        content = re.sub(r'!\[[^\]]*\]\([^\)]*\)', '', content)

        # Remove external email warning
        content = re.sub(
            r'CAUTION:?\s*This email originated from outside.*?(?=\n\n|\n[A-Z]|$)',
            '', content, flags=re.IGNORECASE | re.DOTALL
        )

        # Clean mailto links
        content = re.sub(r'\[([^\]]+@[^\]]+)\]\(mailto:[^)]+\)', r'\1', content)

        # Clean up whitespace
        content = re.sub(r'\n{3,}', '\n\n', content)
        content = re.sub(r'_+\s*', '', content)

        return content.strip()

    def _collect_participants(self, emails: list[Email]) -> list[str]:
        """Collect unique participants from emails."""
        participants = set()
        for email in emails:
            if email.sender_clean and email.sender_clean != "Unknown":
                participants.add(email.sender_clean)
            for name in email.to_names:
                if name:
                    participants.add(name)
        return sorted(list(participants))

    def _determine_tags(self, emails: list[Email]) -> list[str]:
        """Determine appropriate tags for a thread."""
        tags = ["email/thread"]

        has_sent = any(e.is_sent for e in emails)
        has_received = any(not e.is_sent for e in emails)

        if has_sent and has_received:
            tags.append("email/conversation")
        elif has_sent:
            tags.append("email/sent")
        else:
            tags.append("email/received")

        return tags

    def _format_date_wikilink(self, date_obj) -> str:
        """Format date as Obsidian wikilink."""
        if not date_obj:
            return "Unknown"
        try:
            return f"[[{date_obj.strftime('%Y-%m-%d')}]]"
        except Exception:
            return "Unknown"
