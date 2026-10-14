"""MIS e-mail reporting for Diffly.

Every completed comparison produces one MIS mail covering:
  * when the app was used (IST and UTC) and by whom (session, IP, location)
  * which files were uploaded (name, format, size, pages, words)
  * what changed (highlighted words per document, match rate, change rate)

Settings resolve in this order: environment variable -> ``st.secrets`` -> the
built-in default, so the same module works locally and on Streamlit Cloud, and a
deployment can redirect or disable reporting without editing the secrets file.
"""

from __future__ import annotations

import json
import os
import smtplib
import ssl
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from typing import Any

IST = timezone(timedelta(hours=5, minutes=30))

SMTP_HOST = "smtp.gmail.com"
SMTP_TIMEOUT = 25

# Transports are tried in order. 587/STARTTLS goes first because corporate
# networks and proxies frequently drop the implicit-TLS connection on 465,
# while 465 is kept as a fallback for networks that only allow that port.
SMTP_TRANSPORTS = (("starttls", 587), ("ssl", 465))

BRAND = "#005eac"
HEADING = "#00407a"      # dark blue, readable on a light band in every mail client
SUBHEADING = "#3a6184"
HEADER_FILL = "#dce9f6"  # light fill behind table headings
MUTED = "#5b6b7f"        # footer / secondary text, kept dark enough to read

# Non-secret defaults only. GMAIL_APP_PASSWORD is deliberately absent so the
# password is never committed: set it in .streamlit/secrets.toml (git-ignored)
# locally, or in the Secrets box of the Streamlit Cloud app.
_DEFAULTS = {
    "GMAIL_USERNAME": "commissionsmis@gmail.com",
    "MIS_EMAIL_TO": "Ayushi.Roy01@bajajlife.com,bhagyashree.daga@bajajlife.com",
}


def _config(key: str) -> str:
    """Read a setting from the environment, then st.secrets, then the default.

    The environment wins so a deployment can redirect or switch off reporting
    without editing the secrets file - setting MIS_EMAIL_TO empty, for instance,
    leaves reporting unconfigured and nothing is sent.
    """
    value = os.getenv(key)
    if value is not None:
        return value.strip()

    try:
        import streamlit as st

        if key in st.secrets:
            return str(st.secrets[key]).strip()
    except Exception:
        # No secrets file configured, or Streamlit not running - fall through.
        pass

    return _DEFAULTS.get(key, "")


def recipients() -> list[str]:
    raw = _config("MIS_EMAIL_TO")
    return [address.strip() for address in raw.split(",") if address.strip()]


def is_configured() -> bool:
    return bool(_config("GMAIL_USERNAME") and _config("GMAIL_APP_PASSWORD") and recipients())


def _fmt_bytes(size: Any) -> str:
    try:
        size = int(size)
    except (TypeError, ValueError):
        return "-"
    if size <= 0:
        return "-"
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.2f} MB"


def _fmt_int(value: Any) -> str:
    """Indian digit grouping, matching the rest of the app."""
    try:
        number = int(value)
    except (TypeError, ValueError):
        return str(value)

    negative = number < 0
    digits = str(abs(number))
    if len(digits) <= 3:
        grouped = digits
    else:
        tail, head = digits[-3:], digits[:-3]
        chunks = []
        while len(head) > 2:
            chunks.insert(0, head[-2:])
            head = head[:-2]
        if head:
            chunks.insert(0, head)
        grouped = ",".join(chunks + [tail])

    return f"-{grouped}" if negative else grouped


def _rows(pairs: list[tuple[str, str]]) -> str:
    cell_style = "padding:7px 12px;border-bottom:1px solid #e6eef7;font-size:13px"
    cells = []
    for label, value in pairs:
        cells.append(
            f'<tr><td style="{cell_style};color:#475569;white-space:nowrap">{label}</td>'
            f'<td style="{cell_style};color:#0f172a;font-weight:600">{value}</td></tr>'
        )
    return (
        '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;'
        'border:1px solid #e6eef7;border-radius:10px;overflow:hidden">'
        + "".join(cells)
        + "</table>"
    )


def _section(title: str, body: str) -> str:
    return (
        '<div style="margin:0 0 20px 0">'
        '<div style="font-size:12px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;'
        f'color:{BRAND};margin:0 0 8px 2px">{title}</div>{body}</div>'
    )


def build_report(payload: dict[str, Any]) -> dict[str, str]:
    """Build the subject, HTML body and plain-text body from a comparison payload."""
    used_at: datetime = payload.get("used_at") or datetime.now(timezone.utc)
    used_ist = used_at.astimezone(IST)

    doc1 = payload.get("doc1", {})
    doc2 = payload.get("doc2", {})
    summary = payload.get("summary", {})

    changes = summary.get("highlighted_changes", 0)
    largest = max(summary.get("total_words1", 0), summary.get("total_words2", 0))
    change_rate = round((changes / largest) * 100, 2) if largest else 0.0

    subject = (
        f"Doc Compare MIS | {doc1.get('name', 'Document 1')} vs {doc2.get('name', 'Document 2')}"
        f" | {_fmt_int(changes)} changes | {used_ist.strftime('%d-%b-%Y %H:%M')} IST"
    )

    usage = _rows(
        [
            ("Date &amp; time (IST)", used_ist.strftime("%d %b %Y, %I:%M:%S %p")),
            ("Timestamp (UTC)", used_at.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")),
            ("Processing time", f"{payload.get('duration_seconds', 0):.1f} seconds"),
            ("Session ID", payload.get("session_id", "-")),
            ("Client IP", payload.get("client_ip", "unknown")),
            ("Location", payload.get("client_location", "unknown")),
            ("Environment", payload.get("environment", "-")),
        ]
    )

    def _doc_row(label: str, doc: dict[str, Any], words: int, changed: int) -> str:
        cell = "padding:9px 12px;border-bottom:1px solid #e6eef7;font-size:13px"
        pages = doc.get("page_count") or 0
        return (
            f'<tr><td style="{cell};color:#475569">{label}</td>'
            f'<td style="{cell};color:#0f172a;font-weight:600;word-break:break-all">'
            f'{doc.get("name", "-")}</td>'
            f'<td style="{cell};color:#0f172a;text-transform:uppercase">{doc.get("extension", "-")}</td>'
            f'<td style="{cell};color:#0f172a">{_fmt_bytes(doc.get("size_bytes", 0))}</td>'
            f'<td style="{cell};color:#0f172a">{_fmt_int(pages) if pages else "-"}</td>'
            f'<td style="{cell};color:#0f172a">{_fmt_int(words)}</td>'
            f'<td style="{cell};color:#b45309;font-weight:700">{_fmt_int(changed)}</td></tr>'
        )

    headings = ["#", "File uploaded", "Format", "Size", "Pages", "Words", "Highlighted"]
    # Dark text on a light fill rather than white on blue. Light text anywhere in
    # an e-mail is a liability: any client that drops the background - Outlook
    # among them - renders it white on white and the row disappears.
    header_cells = "".join(
        f'<th bgcolor="{HEADER_FILL}" style="padding:9px 12px;text-align:left;font-size:11px;'
        f'letter-spacing:.06em;text-transform:uppercase;color:{HEADING};font-weight:700;'
        f'background-color:{HEADER_FILL};border-bottom:2px solid {BRAND}">{heading}</th>'
        for heading in headings
    )

    files_table = (
        '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;'
        'border:1px solid #e6eef7;border-radius:10px;overflow:hidden">'
        f"<tr>{header_cells}</tr>"
        + _doc_row("1", doc1, summary.get("total_words1", 0), summary.get("diff_words1", 0))
        + _doc_row("2", doc2, summary.get("total_words2", 0), summary.get("diff_words2", 0))
        + "</table>"
    )

    changes_table = _rows(
        [
            ("Total highlighted changes", f'<span style="color:#b45309">{_fmt_int(changes)}</span>'),
            ("Highlighted in document 1", _fmt_int(summary.get("diff_words1", 0))),
            ("Highlighted in document 2", _fmt_int(summary.get("diff_words2", 0))),
            ("Matching words", _fmt_int(summary.get("total_matching", 0))),
            ("Match rate", f'<span style="color:#15803d">{summary.get("match_rate", 0)}%</span>'),
            ("Change rate", f"{change_rate}%"),
        ]
    )

    outputs = _rows(
        [
            ("Highlighted output 1", doc1.get("output_name", "-")),
            ("Highlighted output 2", doc2.get("output_name", "-")),
        ]
    )

    html = f"""<!doctype html>
<html><body style="margin:0;padding:24px;background:#f1f6fc;
font-family:'Segoe UI',Helvetica,Arial,sans-serif">
  <div style="max-width:720px;margin:0 auto;background:#ffffff;border-radius:16px;
  border:1px solid #dbe7f4;overflow:hidden">
    <!-- Solid light band with dark text. A gradient here rendered as white-on-white
         in clients that strip CSS gradients (Outlook among them), which hid the
         heading entirely; dark blue on a light background survives everywhere, and
         stays readable even if the background itself is stripped. -->
    <div bgcolor="#eaf2fa" style="background-color:#eaf2fa;padding:22px 26px;
    border-bottom:3px solid {BRAND}">
      <div style="color:{HEADING};font-size:20px;font-weight:700;letter-spacing:.01em">
        Diffly &mdash; MIS Report</div>
      <div style="color:{SUBHEADING};font-size:13px;margin-top:4px">
        Comparison completed on {used_ist.strftime('%d %b %Y at %I:%M %p')} IST</div>
    </div>
    <div style="padding:24px 26px">
      <div style="margin:0 0 22px 0">
        <span style="display:inline-block;background:#fff7e6;border:1px solid #fcd34d;color:#92400e;
        border-radius:999px;padding:7px 16px;font-size:13px;font-weight:700">
          {_fmt_int(changes)} changes highlighted</span>
        <span style="display:inline-block;background:#ecfdf5;border:1px solid #86efac;color:#166534;
        border-radius:999px;padding:7px 16px;font-size:13px;font-weight:700;margin-left:8px">
          {summary.get('match_rate', 0)}% match</span>
      </div>
      {_section("Usage details", usage)}
      {_section("Files uploaded", files_table)}
      {_section("Changes highlighted", changes_table)}
      {_section("Outputs generated", outputs)}
      <p style="margin:22px 0 0 0;padding-top:14px;border-top:1px solid #e6eef7;color:{MUTED};
      font-size:11px;line-height:1.6">
        Auto-generated by Diffly. The attached JSON carries the same
        record in machine-readable form for MIS consolidation.</p>
    </div>
  </div>
</body></html>"""

    text = (
        "Diffly - MIS Report\n"
        f"Used at : {used_ist.strftime('%d %b %Y, %I:%M:%S %p')} IST\n"
        f"Files   : {doc1.get('name', '-')} ({doc1.get('extension', '-')}) vs "
        f"{doc2.get('name', '-')} ({doc2.get('extension', '-')})\n"
        f"Changes : {_fmt_int(changes)} highlighted "
        f"({_fmt_int(summary.get('diff_words1', 0))} in doc 1, "
        f"{_fmt_int(summary.get('diff_words2', 0))} in doc 2)\n"
        f"Match   : {summary.get('match_rate', 0)}%\n"
        f"Session : {payload.get('session_id', '-')} | IP {payload.get('client_ip', 'unknown')}\n"
    )

    return {"subject": subject, "html": html, "text": text}


def send_mis_email(payload: dict[str, Any]) -> tuple[bool, str]:
    """Send one MIS mail. Returns ``(sent, message)`` and never raises."""
    if not is_configured():
        return False, "MIS mail is not configured (missing Gmail credentials or recipients)."

    username = _config("GMAIL_USERNAME")
    password = _config("GMAIL_APP_PASSWORD")
    to_list = recipients()

    report = build_report(payload)

    message = EmailMessage()
    message["Subject"] = report["subject"]
    message["From"] = f"Diffly <{username}>"
    message["To"] = ", ".join(to_list)
    message.set_content(report["text"])
    message.add_alternative(report["html"], subtype="html")

    used_at = payload.get("used_at") or datetime.now(timezone.utc)
    record = {
        "used_at_ist": used_at.astimezone(IST).isoformat(),
        "used_at_utc": used_at.astimezone(timezone.utc).isoformat(),
        "duration_seconds": round(payload.get("duration_seconds", 0), 2),
        "session_id": payload.get("session_id"),
        "client_ip": payload.get("client_ip"),
        "client_location": payload.get("client_location"),
        "environment": payload.get("environment"),
        "doc1": payload.get("doc1"),
        "doc2": payload.get("doc2"),
        "summary": payload.get("summary"),
    }
    message.add_attachment(
        json.dumps(record, indent=2, default=str).encode("utf-8"),
        maintype="application",
        subtype="json",
        filename=f"mis_{used_at.astimezone(IST).strftime('%Y%m%d_%H%M%S')}.json",
    )

    context = ssl.create_default_context()
    failures: list[str] = []

    for mode, port in SMTP_TRANSPORTS:
        try:
            if mode == "ssl":
                with smtplib.SMTP_SSL(SMTP_HOST, port, context=context, timeout=SMTP_TIMEOUT) as server:
                    server.login(username, password)
                    server.send_message(message)
            else:
                with smtplib.SMTP(SMTP_HOST, port, timeout=SMTP_TIMEOUT) as server:
                    server.ehlo()
                    server.starttls(context=context)
                    server.ehlo()
                    server.login(username, password)
                    server.send_message(message)
        except smtplib.SMTPAuthenticationError:
            # Bad credentials fail the same way on every port - stop here.
            return False, (
                "Gmail rejected the credentials. Check GMAIL_USERNAME and the "
                "16-character app password in your Streamlit secrets."
            )
        except Exception as ex:  # network, DNS, TLS, blocked port...
            failures.append(f"{port}: {ex}")
            continue

        return True, f"MIS report sent to {', '.join(to_list)}"

    return False, "MIS mail could not be sent (" + "; ".join(failures) + ")"
