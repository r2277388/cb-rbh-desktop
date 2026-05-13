from __future__ import annotations

from collections.abc import Iterable
import os
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
OL_FOLDER_SENT_MAIL = 5


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


def _account_text(account) -> str:
    parts = []
    for attribute_name in ("DisplayName", "SmtpAddress"):
        try:
            value = getattr(account, attribute_name)
        except Exception:
            value = None
        if value and value not in parts:
            parts.append(str(value))
    try:
        store_name = account.DeliveryStore.DisplayName
    except Exception:
        store_name = None
    if store_name and store_name not in parts:
        parts.append(f"store={store_name}")
    return " / ".join(parts) if parts else str(account)


def _iter_accounts(session):
    accounts = session.Accounts
    for index in range(1, accounts.Count + 1):
        yield accounts.Item(index)


def _default_account(session):
    try:
        default_store_id = session.DefaultStore.StoreID
    except Exception:
        default_store_id = None

    fallback = None
    for account in _iter_accounts(session):
        if fallback is None:
            fallback = account
        if default_store_id:
            try:
                if account.DeliveryStore.StoreID == default_store_id:
                    return account
            except Exception:
                pass
    return fallback


def _find_account(session, account_hint: str | None):
    accounts = list(_iter_accounts(session))
    if not account_hint:
        return _default_account(session), accounts

    normalized_hint = account_hint.lower().strip()
    for account in accounts:
        searchable = _account_text(account).lower()
        if normalized_hint in searchable:
            return account, accounts

    visible_accounts = "\n  ".join(_account_text(account) for account in accounts)
    raise RuntimeError(
        f"Outlook account matching '{account_hint}' was not found. "
        f"Visible account(s):\n  {visible_accounts}"
    )


def _configure_outlook_account(mail_item, account, accounts) -> None:
    if account is None:
        return

    try:
        mail_item.SendUsingAccount = account
    except Exception as exc:
        print(f"Outlook account warning: could not set sending account: {exc}")

    try:
        mail_item.DeleteAfterSubmit = False
    except Exception:
        pass

    try:
        mail_item.SaveSentMessageFolder = account.DeliveryStore.GetDefaultFolder(
            OL_FOLDER_SENT_MAIL
        )
    except Exception:
        pass

    if len(accounts) > 1:
        account_lines = "\n    ".join(_account_text(candidate) for candidate in accounts)
        print(
            "Outlook account check: multiple accounts are visible to classic Outlook.\n"
            f"  Using: {_account_text(account)}\n"
            "  Sent copy target: that account's Sent Items folder.\n"
            "  To force a different account, set OUTLOOK_ACCOUNT to part of its name or email.\n"
            "  Visible accounts:\n"
            f"    {account_lines}"
        )
    else:
        print(f"Outlook account check: using {_account_text(account)}")


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
    account_hint: str | None = None,
):
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        session = outlook.Session
        selected_account, accounts = _find_account(
            session,
            account_hint or os.environ.get("OUTLOOK_ACCOUNT"),
        )
        mail_item = outlook.CreateItem(OL_MAIL_ITEM)
        _configure_outlook_account(mail_item, selected_account, accounts)
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
    account_hint: str | None = None,
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
        account_hint=account_hint,
    )
    if display_before_send:
        mail_item.Save()
        mail_item.Display(modal_display)
        pythoncom.CoUninitialize()
        return mail_item

    mail_item.Send()
    pythoncom.CoUninitialize()
    return mail_item
