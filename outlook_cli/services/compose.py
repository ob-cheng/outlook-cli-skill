"""Email compose and send service."""

from pathlib import Path


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
        html: bool = False,
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

            # Subject and body
            mail.Subject = subject
            if html:
                mail.HTMLBody = body
            else:
                mail.Body = body

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
        html: bool = False,
        send_immediately: bool = True,
    ) -> tuple[bool, str]:
        """Reply to an existing email.

        Args:
            message_id: The EntryID of the message to reply to
            body: Reply message body
            reply_all: If True, reply to all recipients; if False, reply to sender only
            attachments: List of file paths to attach
            html: If True, body is HTML; otherwise plain text
            send_immediately: If True, send now; if False, save to Drafts

        Returns:
            tuple: (success: bool, message_id_or_error: str)
        """
        try:
            original = self.namespace.GetItemFromID(message_id)

            if reply_all:
                reply = original.ReplyAll()
            else:
                reply = original.Reply()

            # Add body (prepend to existing quoted text)
            if html:
                reply.HTMLBody = body + "<br><br>" + reply.HTMLBody
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
        html: bool = False,
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

            # Add forwarding message
            if body:
                if html:
                    forward.HTMLBody = body + "<br><br>" + forward.HTMLBody
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

    def create_draft(
        self,
        to: list[str] | None = None,
        subject: str = "",
        body: str = "",
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        attachments: list[str | Path] | None = None,
        html: bool = False,
    ) -> tuple[bool, str]:
        """Create a draft email without sending.

        Args:
            to: List of recipient email addresses
            subject: Email subject
            body: Email body
            cc: List of CC email addresses
            bcc: List of BCC email addresses
            attachments: List of file paths to attach
            html: If True, body is HTML; otherwise plain text

        Returns:
            tuple: (success: bool, message_id_or_error: str)
        """
        return self.send_email(
            to=to or [],
            subject=subject,
            body=body,
            cc=cc,
            bcc=bcc,
            attachments=attachments,
            html=html,
            send_immediately=False,
        )
