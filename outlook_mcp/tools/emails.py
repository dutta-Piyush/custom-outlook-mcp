"""Outlook read tools: list/search/read emails, list folders and attachments."""

from __future__ import annotations

import logging
from urllib.parse import urlencode

from outlook_mcp.app import mcp
from outlook_mcp.client import GraphError, get_client
from outlook_mcp.utils import sanitize_odata_param, strip_html, truncate

logger = logging.getLogger(__name__)


@mcp.tool()
def list_emails(
    folder: str = "Inbox",
    top: int = 10,
    skip: int = 0,
    filter: str = "",
    search: str = "",
    orderby: str = "",
    select: str = "",
) -> str:
    """List emails from a mail folder with filtering, searching and pagination.

    Args:
        folder: Mail folder name or ID (default: Inbox). Well-known names:
                Inbox, Drafts, SentItems, DeletedItems, JunkEmail, Archive.
        top: Number of emails to return (default 10, max 50)
        skip: Emails to skip for pagination (default 0)
        filter: OData $filter expression, e.g. 'isRead eq false',
                'importance eq \'high\'', 'hasAttachments eq true',
                'receivedDateTime ge 2026-01-01T00:00:00Z',
                'contains(subject,\'report\')'
        search: KQL search query (cannot combine with orderby), e.g.
                'from:john@example.com', 'subject:budget', 'hasAttachments:true'
        orderby: Sort order (cannot combine with search), e.g.
                 'receivedDateTime desc', 'subject asc'
        select: Comma-separated fields to return. Defaults to
                id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview.
                Add: importance,flag,categories,conversationId,toRecipients
    """
    c = get_client()
    if top <= 0 or top > 50:
        top = 10

    # Sanitize caller-supplied OData parameters before embedding in the query
    try:
        if filter:
            filter = sanitize_odata_param(filter, "filter")
        if search:
            search = sanitize_odata_param(search, "search")
        if orderby:
            orderby = sanitize_odata_param(orderby, "orderby")
        if select:
            select = sanitize_odata_param(select, "select")
    except ValueError as exc:
        return f"ERROR: {exc}"

    select_fields = select or "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview"

    params: dict = {
        "$top": top,
        "$select": select_fields,
    }
    if skip > 0:
        params["$skip"] = skip
    if filter:
        params["$filter"] = filter
    if search:
        params["$search"] = f'"{search}"'
    else:
        params["$orderby"] = orderby or "receivedDateTime desc"

    path = f"/me/mailFolders/{folder}/messages?{urlencode(params)}"
    logger.info("list_emails: folder=%s top=%d skip=%d", folder, top, skip)
    try:
        data = c.get(path)
    except GraphError as e:
        logger.error("list_emails failed: %s", e)
        return f"ERROR: {e}"

    emails = data.get("value", [])
    next_link = data.get("@odata.nextLink")

    lines = [f"{len(emails)} emails from {folder}:\n"]
    for i, email in enumerate(emails, 1):
        from_ea = ((email.get("from") or {}).get("emailAddress") or {})
        sender = f"{from_ea.get('name', '')} <{from_ea.get('address', '')}>"
        read = "[READ]" if email.get("isRead") else "[UNREAD]"
        attach = " [HAS ATTACHMENT]" if email.get("hasAttachments") else ""
        imp = " [HIGH IMPORTANCE]" if email.get("importance") == "high" else ""
        preview = truncate(str(email.get("bodyPreview", "")), 100)
        lines.append(
            f"[{i}] {read} {email.get('subject', '')}{attach}{imp}\n"
            f"    From: {sender}\n"
            f"    Date: {email.get('receivedDateTime', '')}\n"
            f"    ID: {email.get('id', '')}\n"
            f"    Preview: {preview}\n"
        )

    if next_link:
        lines.append(f"More results available. Use skip={skip + top} to get next page.")

    return "\n".join(lines)


@mcp.tool()
def read_email(id: str, format: str = "full") -> str:
    """Read a specific email by ID.

    Args:
        id: The email message ID
        format: 'full' (default) — body + metadata;
                'headers' — internet message headers for debugging/tracing
    """
    c = get_client()
    if not id:
        return "ERROR: email id is required"

    if format == "headers":
        select = (
            "id,subject,from,toRecipients,ccRecipients,bccRecipients,"
            "receivedDateTime,sentDateTime,importance,isRead,hasAttachments,"
            "flag,categories,conversationId,internetMessageHeaders"
        )
    else:
        select = (
            "id,subject,from,toRecipients,ccRecipients,"
            "receivedDateTime,body,hasAttachments,importance,isRead,"
            "flag,categories,conversationId,replyTo"
        )

    logger.info("read_email: id=%s format=%s", id, format)
    try:
        email = c.get(f"/me/messages/{id}?$select={select}")
    except GraphError as e:
        logger.error("read_email failed: %s", e)
        return f"ERROR: {e}"

    lines = [
        f"Subject: {email.get('subject', '')}",
        "─" * 40,
    ]

    from_ea = ((email.get("from") or {}).get("emailAddress") or {})
    lines.append(f"From: {from_ea.get('name', '')} <{from_ea.get('address', '')}>")

    to_list = [
        f"{(r.get('emailAddress') or {}).get('name', '')} <{(r.get('emailAddress') or {}).get('address', '')}>"
        for r in (email.get("toRecipients") or [])
    ]
    if to_list:
        lines.append(f"To: {', '.join(to_list)}")

    cc_list = [
        f"{(r.get('emailAddress') or {}).get('name', '')} <{(r.get('emailAddress') or {}).get('address', '')}>"
        for r in (email.get("ccRecipients") or [])
    ]
    if cc_list:
        lines.append(f"CC: {', '.join(cc_list)}")

    lines.append(f"Date: {email.get('receivedDateTime', '')}")
    lines.append(f"Importance: {email.get('importance', 'normal')}")
    lines.append(f"Has Attachments: {email.get('hasAttachments', False)}")

    if email.get("conversationId"):
        lines.append(f"Conversation ID: {email['conversationId']}")
    if email.get("categories"):
        lines.append(f"Categories: {', '.join(email['categories'])}")

    lines.append("─" * 40)
    lines.append("")

    if format == "headers":
        for h in email.get("internetMessageHeaders") or []:
            lines.append(f"{h.get('name', '')}: {h.get('value', '')}")
    else:
        body = email.get("body") or {}
        content = body.get("content", "")
        if body.get("contentType") == "html":
            content = strip_html(content)
        lines.append(content)

    return "\n".join(lines)


@mcp.tool()
def search_emails(query: str, top: int = 10, folder: str = "") -> str:
    """Search emails across the mailbox or a specific folder using KQL.

    Args:
        query: KQL search query. Supports:
               plain text, 'from:sender@email.com', 'to:recipient',
               'subject:keyword', 'hasAttachments:true',
               'received>=2026-01-01', 'importance:high',
               AND/OR operators and (grouping)
        top: Number of results to return (default 10, max 50)
        folder: Limit to a specific folder (default: all folders).
                Use well-known name (Inbox, SentItems, etc.) or folder ID.
    """
    c = get_client()
    if not query:
        return "ERROR: search query is required"
    if top <= 0 or top > 50:
        top = 10

    try:
        query = sanitize_odata_param(query, "search query")
    except ValueError as exc:
        return f"ERROR: {exc}"

    params = urlencode({
        "$search": f'"{query}"',
        "$top": top,
        "$select": "id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments,importance",
    })

    base = f"/me/mailFolders/{folder}/messages" if folder else "/me/messages"
    logger.info("search_emails: query=%r top=%d folder=%s", query, top, folder or "all")
    try:
        data = c.get(f"{base}?{params}")
    except GraphError as e:
        logger.error("search_emails failed: %s", e)
        return f"ERROR: {e}"

    emails = data.get("value", [])
    scope = folder or "all folders"
    lines = [f"Search \"{query}\" in {scope} — {len(emails)} result(s):\n"]

    for i, email in enumerate(emails, 1):
        from_ea = ((email.get("from") or {}).get("emailAddress") or {})
        attach = " [HAS ATTACHMENT]" if email.get("hasAttachments") else ""
        preview = truncate(str(email.get("bodyPreview", "")), 120)
        lines.append(
            f"[{i}] {email.get('subject', '')}{attach}\n"
            f"    From: {from_ea.get('address', '')} | Date: {email.get('receivedDateTime', '')}\n"
            f"    ID: {email.get('id', '')}\n"
            f"    {preview}\n"
        )

    return "\n".join(lines)


@mcp.tool()
def list_folders(include_children: bool = False) -> str:
    """List all mail folders with unread/total counts.

    Args:
        include_children: If True, also lists sub-folders (default False)
    """
    c = get_client()
    try:
        data = c.get(
            "/me/mailFolders?$top=100"
            "&$select=id,displayName,totalItemCount,unreadItemCount,childFolderCount"
        )
    except GraphError as e:
        return f"ERROR: {e}"

    folders = data.get("value", [])
    lines = [f"Mail Folders ({len(folders)}):\n"]

    for folder in folders:
        child_count = int(folder.get("childFolderCount") or 0)
        child_note = f" [{child_count} subfolders]" if child_count > 0 else ""
        lines.append(
            f"  {folder.get('displayName', '')} "
            f"(unread: {folder.get('unreadItemCount', 0)} / total: {folder.get('totalItemCount', 0)})"
            f"{child_note}\n"
            f"    ID: {folder.get('id', '')}"
        )

        if include_children and child_count > 0:
            try:
                child_data = c.get(
                    f"/me/mailFolders/{folder['id']}/childFolders"
                    "?$select=id,displayName,totalItemCount,unreadItemCount"
                )
                for child in child_data.get("value", []):
                    lines.append(
                        f"    └─ {child.get('displayName', '')} "
                        f"(unread: {child.get('unreadItemCount', 0)} / total: {child.get('totalItemCount', 0)})\n"
                        f"       ID: {child.get('id', '')}"
                    )
            except GraphError:
                pass  # skip subfolders that fail

    return "\n".join(lines)


@mcp.tool()
def list_attachments(id: str) -> str:
    """List all attachments on a specific email.

    Args:
        id: The email message ID
    """
    c = get_client()
    if not id:
        return "ERROR: email id is required"

    try:
        data = c.get(f"/me/messages/{id}/attachments?$select=id,name,contentType,size")
    except GraphError as e:
        return f"ERROR: {e}"

    attachments = data.get("value", [])
    lines = [f"Attachments ({len(attachments)}):\n"]
    for i, att in enumerate(attachments, 1):
        size_kb = (att.get("size") or 0) / 1024
        lines.append(
            f"  [{i}] {att.get('name', '')} ({size_kb:.1f} KB, {att.get('contentType', '')})\n"
            f"      ID: {att.get('id', '')}"
        )

    return "\n".join(lines)
