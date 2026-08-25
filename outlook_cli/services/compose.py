"""Email compose and send service."""

import re
from pathlib import Path

_BODY_TAG_RE = re.compile(r"<body[^>]*>", re.IGNORECASE)
_HTML_TAG_RE = re.compile(r"<[a-zA-Z][^>]*>")
_SPAN_STYLE_RE = re.compile(r"<span[^>]*\bstyle=(['\"])(.*?)\1", re.IGNORECASE | re.DOTALL)


def _as_html_fragment(body: str) -> str:
    """Turn a plain-text body into HTML that still shows line breaks.

    HTML collapses bare newlines, so a caller passing plain multi-line text
    (the common case, since html now defaults to True) would otherwise render
    as one run-on paragraph. Bodies that already contain markup are left
    untouched.
    """
    if _HTML_TAG_RE.search(body):
        return body
    return body.replace("\n", "<br>\n")


def _prepend_html(existing_html: str, new_content: str) -> str:
    """Insert new_content just inside <body>, styled to match the signature.

    existing_html (Outlook's own reply/forward/signature markup) is a full
    document, but Word doesn't set font via CSS on <body> — every run of
    text carries its own inline <span style="font-family:...;font-size:...">.
    Content dropped in unstyled falls back to Outlook's compose-window
    default font, not the signature's font. Reusing the signature's own
    first span style keeps the two visually consistent.
    """
    match = _BODY_TAG_RE.search(existing_html)
    if not match:
        return new_content + existing_html
    insert_at = match.end()

    style_match = _SPAN_STYLE_RE.search(existing_html)
    if style_match:
        quote, style = style_match.group(1), style_match.group(2)
        new_content = (
            f"<p class=MsoNormal><span style={quote}{style}{quote}>"
            f"{new_content}</span></p>"
        )

    return existing_html[:insert_at] + new_content + existing_html[insert_at:]


class ComposeService:
    """Service for composing, sending, replying, and drafting emails."""

    def __init__(self, namespace):
        """Initialize with Outlook MAPI namespace."""
        self.namespace = namespace

    def send_email(
        self,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        attachments: list[str | Path] | None = None,
        html: bool = True,
        send_immediately: bool = True,
    ) -> tuple[bool, str]:
        """Send a new email.

        Args:
            to: List of recipient email addresses
            subject: Email subject
            body: Email body
            cc: List of CC email addresses
            bcc: List of BCC email addresses
            attachments: List of file paths to attach
            html: If True, body is HTML; otherwise plain text
            send_immediately: If True, send now; if False, save to Drafts

        Returns:
            tuple: (success: bool, message_id_or_error: str)
        """
        try:
            outlook = self.namespace.Application
            mail = outlook.CreateItem(0)  # 0 = olMailItem

            # Recipients
            mail.To = "; ".join(to)
            if cc:
                mail.CC = "; ".join(cc)
            if bcc:
                mail.BCC = "; ".join(bcc)
            # Setting To/CC/BCC as a string leaves recipients unresolved until
            # Outlook checks them, which shows as an error in the compose UI
            # even for valid addresses. ResolveAll forces that check now.
            mail.Recipients.ResolveAll()

            # Subject and body
            mail.Subject = subject
            # Touching GetInspector forces Outlook to auto-insert the default
            # "new message" signature into Body/HTMLBody, same as opening a
            # compose window in the UI — CreateItem() alone never does this.
            mail.GetInspector
            if html:
                mail.HTMLBody = _prepend_html(
                    mail.HTMLBody, _as_html_fragment(body) + "<br><br>"
                )
            else:
                mail.Body = body + "\n\n" + mail.Body

            # Attachments
            if attachments:
                for file_path in attachments:
                    file_path = Path(file_path)
                    if file_path.exists():
                        mail.Attachments.Add(str(file_path.absolute()))

            # Send or save as draft
            if send_immediately:
                mail.Send()
                return True, "Email sent successfully"
            else:
                mail.Save()
                return True, f"Email saved to Drafts (ID: {mail.EntryID})"

        except Exception as e:
            return False, f"Failed to send email: {e}"

    def reply(
        self,
        message_id: str,
        body: str,
        reply_all: bool = False,
        attachments: list[str | Path] | None = None,
        html: bool = True,
        send_immediately: bool = True,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
    ) -> tuple[bool, str]:
        """Reply to an existing email.

        Args:
            message_id: The EntryID of the message to reply to
            body: Reply message body
            reply_all: If True, reply to all recipients; if False, reply to sender only
            attachments: List of file paths to attach
            html: If True, body is HTML; otherwise plain text
            send_immediately: If True, send now; if False, save to Drafts
            cc: Additional CC recipients to add beyond thread participants
            bcc: Additional BCC recipients to add

        Returns:
            tuple: (success: bool, message_id_or_error: str)
        """
        try:
            original = self.namespace.GetItemFromID(message_id)

            if reply_all:
                reply = original.ReplyAll()
            else:
                reply = original.Reply()

            # Add extra CC/BCC recipients
            if cc:
                for addr in cc:
                    recipient = reply.Recipients.Add(addr)
                    recipient.Type = 2  # olCC
            if bcc:
                for addr in bcc:
                    recipient = reply.Recipients.Add(addr)
                    recipient.Type = 3  # olBCC
            if cc or bcc:
                reply.Recipients.ResolveAll()

            # Add body (prepend to existing quoted text)
            if html:
                reply.HTMLBody = _prepend_html(
                    reply.HTMLBody, _as_html_fragment(body) + "<br><br>"
                )
            else:
                reply.Body = body + "\n\n" + reply.Body

            # Attachments
            if attachments:
                for file_path in attachments:
                    file_path = Path(file_path)
                    if file_path.exists():
                        reply.Attachments.Add(str(file_path.absolute()))

            # Send or save as draft
            if send_immediately:
                reply.Send()
                return True, "Reply sent successfully"
            else:
                reply.Save()
                return True, f"Reply saved to Drafts (ID: {reply.EntryID})"

        except Exception as e:
            return False, f"Failed to reply: {e}"

    def forward(
        self,
        message_id: str,
        to: list[str],
        body: str | None = None,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        attachments: list[str | Path] | None = None,
        html: bool = True,
        send_immediately: bool = True,
    ) -> tuple[bool, str]:
        """Forward an existing email.

        Args:
            message_id: The EntryID of the message to forward
            to: List of recipient email addresses
            body: Optional message to prepend
            cc: List of CC email addresses
            bcc: List of BCC email addresses
            attachments: Additional file paths to attach
            html: If True, body is HTML; otherwise plain text
            send_immediately: If True, send now; if False, save to Drafts

        Returns:
            tuple: (success: bool, message_id_or_error: str)
        """
        try:
            original = self.namespace.GetItemFromID(message_id)
            forward = original.Forward()

            # Recipients
            forward.To = "; ".join(to)
            if cc:
                forward.CC = "; ".join(cc)
            if bcc:
                forward.BCC = "; ".join(bcc)
            forward.Recipients.ResolveAll()

            # Add forwarding message
            if body:
                if html:
                    forward.HTMLBody = _prepend_html(
                        forward.HTMLBody, _as_html_fragment(body) + "<br><br>"
                    )
                else:
                    forward.Body = body + "\n\n" + forward.Body

            # Additional attachments
            if attachments:
                for file_path in attachments:
                    file_path = Path(file_path)
                    if file_path.exists():
                        forward.Attachments.Add(str(file_path.absolute()))

            # Send or save as draft
            if send_immediately:
                forward.Send()
                return True, "Email forwarded successfully"
            else:
                forward.Save()
                return True, f"Forward saved to Drafts (ID: {forward.EntryID})"

        except Exception as e:
            return False, f"Failed to forward: {e}"
