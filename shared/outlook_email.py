from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

try:
    import pythoncom
    import win32com.client
except ModuleNotFoundError as exc:
    missing_module = exc.name or "pywin32"
    raise RuntimeError(
        "Outlook automation requires pywin32. "
        f"Missing module: {missing_module}."
    ) from exc


OL_MAIL_ITEM = 0


def _normalize_recipients(recipients: str | Iterable[str] | None) -> str:
    if recipients is None:
        return ""
    if isinstance(recipients, str):
        values = [recipients]
    else:
        values = list(recipients)
    cleaned = [value.strip() for value in values if value and value.strip()]
    return "; ".join(cleaned)


def _resolve_recipients(mail_item) -> None:
    recipients = mail_item.Recipients
    if recipients.Count == 0:
        return
    if not recipients.ResolveAll():
        unresolved = []
        for index in range(1, recipients.Count + 1):
            recipient = recipients.Item(index)
            if not recipient.Resolved:
                unresolved.append(recipient.Name)
        if unresolved:
            raise RuntimeError(
                "Outlook could not resolve recipient(s): "
                + "; ".join(unresolved)
            )


def create_outlook_mail(
    *,
    to: str | Iterable[str],
    subject: str,
    body: str | None = None,
    html_body: str | None = None,
    cc: str | Iterable[str] | None = None,
    bcc: str | Iterable[str] | None = None,
    attachments: Iterable[str | Path] | None = None,
    resolve_recipients: bool = True,
):
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail_item = outlook.CreateItem(OL_MAIL_ITEM)
        mail_item.To = _normalize_recipients(to)
        mail_item.CC = _normalize_recipients(cc)
        mail_item.BCC = _normalize_recipients(bcc)
        mail_item.Subject = subject

        if html_body is not None:
            mail_item.HTMLBody = html_body
        elif body is not None:
            mail_item.Body = body

        if attachments:
            for attachment in attachments:
                mail_item.Attachments.Add(str(Path(attachment)))

        if resolve_recipients:
            _resolve_recipients(mail_item)

        return mail_item
    except Exception:
        pythoncom.CoUninitialize()
        raise


def send_outlook_mail(
    *,
    to: str | Iterable[str],
    subject: str,
    body: str | None = None,
    html_body: str | None = None,
    cc: str | Iterable[str] | None = None,
    bcc: str | Iterable[str] | None = None,
    attachments: Iterable[str | Path] | None = None,
    display_before_send: bool = False,
    modal_display: bool = False,
    resolve_recipients: bool = True,
):
    mail_item = create_outlook_mail(
        to=to,
        subject=subject,
        body=body,
        html_body=html_body,
        cc=cc,
        bcc=bcc,
        attachments=attachments,
        resolve_recipients=resolve_recipients,
    )
    if display_before_send:
        mail_item.Save()
        mail_item.Display(modal_display)
        pythoncom.CoUninitialize()
        return mail_item

    mail_item.Send()
    pythoncom.CoUninitialize()
    return mail_item
