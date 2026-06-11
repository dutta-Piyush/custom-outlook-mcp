"""Outlook write/action tools: send, reply, move emails."""

from __future__ import annotations

import logging

from outlook_mcp.app import mcp
from outlook_mcp.client import GraphError, get_client
from outlook_mcp.utils import build_recipients

logger = logging.getLogger(__name__)


@mcp.tool()
def send_email(
    to: str,
    subject: str,
    body: str = "",
    body_type: str = "Text",
    cc: str = "",
    bcc: str = "",
    importance: str = "normal",
) -> str:
    """Send a new email via Microsoft Graph API.

    Args:
        to: Comma-separated recipient email addresses,
            e.g. 'alice@example.com, Bob <bob@example.com>'
        subject: Email subject line
        body: Email body content
        body_type: 'Text' (default) or 'HTML'
        cc: Comma-separated CC addresses (optional)
        bcc: Comma-separated BCC addresses (optional)
        importance: 'low', 'normal' (default), or 'high'
    """
    c = get_client()
    if not to or not subject:
        return "ERROR: 'to' and 'subject' are required"

    message: dict = {
        "subject": subject,
        "body": {
            "contentType": body_type or "Text",
            "content": body,
        },
        "toRecipients": build_recipients(to),
        "importance": importance or "normal",
    }
    if cc:
        message["ccRecipients"] = build_recipients(cc)
    if bcc:
        message["bccRecipients"] = build_recipients(bcc)

    logger.info("send_email: to=%s subject=%r importance=%s", to, subject, importance)
    try:
        c.post("/me/sendMail", {"message": message})
    except GraphError as e:
        logger.error("send_email failed: %s", e)
        return f"ERROR: {e}"

    logger.info("send_email: sent successfully to %s", to)
    return f"Email sent to: {to}\nSubject: {subject}\nImportance: {importance}"


@mcp.tool()
def reply_email(id: str, comment: str, reply_all: bool = False) -> str:
    """Reply to an email.

    Args:
        id: The email message ID to reply to
        comment: Reply message body text
        reply_all: If True, replies to all recipients. Default False (sender only)
    """
    c = get_client()
    if not id or not comment:
        return "ERROR: 'id' and 'comment' are required"

    action = "replyAll" if reply_all else "reply"
    logger.info("reply_email: id=%s reply_all=%s", id, reply_all)
    try:
        c.post(f"/me/messages/{id}/{action}", {"comment": comment})
    except GraphError as e:
        logger.error("reply_email failed: %s", e)
        return f"ERROR: {e}"

    label = "Reply-all" if reply_all else "Reply"
    return f"{label} sent for message {id}"


@mcp.tool()
def move_email(id: str, folder_id: str) -> str:
    """Move an email to a different folder.

    Args:
        id: The email message ID to move
        folder_id: Destination folder ID or well-known name:
                   Inbox, Drafts, SentItems, DeletedItems, Archive, JunkEmail
    """
    c = get_client()
    if not id or not folder_id:
        return "ERROR: 'id' and 'folder_id' are required"

    logger.info("move_email: id=%s -> folder=%s", id, folder_id)
    try:
        c.post(f"/me/messages/{id}/move", {"destinationId": folder_id})
    except GraphError as e:
        logger.error("move_email failed: %s", e)
        return f"ERROR: {e}"

    logger.info("move_email: moved %s to %s", id, folder_id)
    return f"Email {id} moved to folder: {folder_id}"
