"""Document Comparison Studio - Streamlit front-end.

Runs on the current comparison engine (``comparison_core``), the same one the
FastAPI/React build uses, so the Streamlit deployment and the API stay in step.

The side-by-side preview lives in ``preview_viewer`` - Streamlit cannot
synchronise scrolling between two of its own containers, so both documents and
their controls are rendered into one embedded HTML document.

Each completed comparison is e-mailed as an MIS report (see ``mis_mailer``).
That is a back-office concern: it is deliberately invisible in the interface,
and a failure to send is logged rather than shown to the person comparing.
"""

from __future__ import annotations

import base64
import json
import logging
import os
import platform
import socket
import time
import uuid
from datetime import datetime, timedelta, timezone
from functools import lru_cache
from io import BytesIO
from pathlib import Path
from urllib import request as urllib_request
from urllib.error import HTTPError, URLError
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st

import mis_mailer
from comparison_core import ALLOWED_EXTENSIONS, compare_documents_full
from preview_viewer import build_viewer_html, viewer_data_url
from usage_storage import get_usage_log_path

logger = logging.getLogger(__name__)

APP_DIR = Path(__file__).resolve().parent
USAGE_LOG_FILE = get_usage_log_path(__file__)
IST = timezone(timedelta(hours=5, minutes=30))
VIEWER_HEIGHT = 880

MIME_BY_EXT = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}

st.set_page_config(
    page_title="Document Comparison Studio",
    page_icon=str(APP_DIR / "logo.png") if (APP_DIR / "logo.png").exists() else ":material/difference:",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def format_indian_number(value) -> str:
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


def _safe_get_headers() -> dict:
    context = getattr(st, "context", None)
    if not context:
        return {}
    headers = getattr(context, "headers", None)
    return dict(headers) if headers else {}


def _extract_client_ip(headers: dict) -> str:
    for key in ["x-forwarded-for", "x-real-ip", "cf-connecting-ip", "x-client-ip"]:
        value = headers.get(key) or headers.get(key.title())
        if value:
            return value.split(",")[0].strip()
    return "unknown"


@lru_cache(maxsize=256)
def _resolve_location_from_ip(ip_address: str) -> tuple:
    """Returns (country, region, city, source). Cached - one lookup per IP."""
    if not ip_address or ip_address == "unknown":
        return ("unknown", "unknown", "unknown", "none")

    if ip_address.startswith(("10.", "192.168.", "172.", "127.")):
        return ("private-network", "private-network", "private-network", "private-ip")

    try:
        url = f"https://ipapi.co/{ip_address}/json/"
        with urllib_request.urlopen(url, timeout=2) as response:
            payload = json.loads(response.read().decode("utf-8"))
        return (
            payload.get("country_name") or "unknown",
            payload.get("region") or "unknown",
            payload.get("city") or "unknown",
            "ipapi.co",
        )
    except (URLError, HTTPError, TimeoutError, json.JSONDecodeError, OSError):
        return ("unknown", "unknown", "unknown", "lookup-failed")


def _session_id() -> str:
    if "usage_session_id" not in st.session_state:
        st.session_state.usage_session_id = str(uuid.uuid4())
    return st.session_state.usage_session_id


def _environment_label() -> str:
    if os.getenv("STREAMLIT_RUNTIME_ENV") or os.getenv("HOSTNAME", "").startswith("streamlit"):
        return "Streamlit Cloud"
    try:
        return f"{platform.system()} / {socket.gethostname()}"
    except Exception:
        return platform.system() or "unknown"


def append_usage_event(event: dict) -> None:
    try:
        USAGE_LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(USAGE_LOG_FILE, "a", encoding="utf-8") as handle:
            handle.write(json.dumps(event, ensure_ascii=False) + "\n")
    except OSError:
        pass


@st.cache_data(show_spinner=False)
def _logo_data_uri() -> str:
    logo_path = APP_DIR / "logo.png"
    if not logo_path.exists():
        return ""
    try:
        return base64.b64encode(logo_path.read_bytes()).decode("utf-8")
    except OSError:
        return ""


# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------

st.markdown(
    """
<style>
    :root {
        --brand: #005eac;
        --ink: #0f172a;
        --muted: #64748b;
        --faint: #94a3b8;
        --line: #dde8f4;
        --warn: #b45309;
        --ok: #15803d;
    }

    html, body, [class*="css"] { font-family: "Inter", "Segoe UI", system-ui, Arial, sans-serif; }

    .stApp { background: #f6f9fd; }

    /* The viewer needs every pixel of width, so the page runs nearly full-bleed
       and there is no sidebar competing for room. */
    .block-container { padding: 1.9rem 1.6rem 2.6rem; max-width: 100%; }
    header[data-testid="stHeader"] { background: transparent; height: 0; }
    #MainMenu, footer { visibility: hidden; }
    section[data-testid="stSidebar"], div[data-testid="stSidebarCollapsedControl"] { display: none; }

    /* ---------- header ---------- */
    .hdr { display: flex; align-items: center; gap: 14px; margin: 0 0 22px; }
    .hdr img { width: 46px; height: 46px; object-fit: contain; flex: 0 0 auto; }
    .hdr h1 {
        color: var(--ink); margin: 0; font-size: 1.42rem; font-weight: 700; letter-spacing: -.02em;
    }
    .hdr p { color: var(--muted); margin: 2px 0 0; font-size: .86rem; }

    .step {
        font-size: .95rem; font-weight: 700; color: var(--ink);
        margin: 0 0 10px; letter-spacing: -.01em;
    }
    .step small { font-weight: 500; color: var(--faint); font-size: .82rem; margin-left: 8px; }

    /* ---------- metrics ---------- */
    .mrow { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 10px; margin: 0 0 18px; }
    .mcard {
        background: #fff; border: 1px solid var(--line); border-radius: 10px; padding: 12px 14px;
    }
    .mcard .l {
        font-size: .68rem; text-transform: uppercase; letter-spacing: .07em;
        color: var(--faint); font-weight: 700;
    }
    .mcard .v {
        font-size: 1.45rem; font-weight: 740; color: var(--ink);
        line-height: 1.2; margin-top: 2px; letter-spacing: -.025em;
    }
    .mcard .s { font-size: .72rem; color: var(--muted);
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .mcard.warn .v { color: var(--warn); }
    .mcard.ok .v { color: var(--ok); }

    /* ---------- legend ---------- */
    .legend {
        display: flex; gap: 16px; flex-wrap: wrap; align-items: center;
        font-size: .78rem; color: var(--muted); margin: 0 0 9px;
    }
    .sw { display: inline-block; width: 12px; height: 12px; border-radius: 3px;
        vertical-align: -2px; margin-right: 6px; border: 1px solid rgba(0,0,0,.14); }
    .kbd {
        font-family: ui-monospace, Menlo, Consolas, monospace; font-size: .69rem;
        border: 1px solid var(--line); border-bottom-width: 2px; border-radius: 4px;
        padding: 1px 5px; background: #fff; color: var(--muted); margin: 0 1px;
    }

    .empty {
        border: 1px dashed #c9dced; border-radius: 10px; background: #fbfdff;
        padding: 26px; text-align: center; color: var(--muted); font-size: .86rem;
    }

    /* ---------- widgets ---------- */
    div[data-testid="stFileUploader"] section {
        border: 1.5px dashed #a9cdec !important; border-radius: 10px !important;
        background: #fff !important; padding: 10px 12px !important;
    }
    div[data-testid="stFileUploader"] section:hover { border-color: var(--brand) !important; }

    .stButton > button, .stDownloadButton > button { border-radius: 9px; font-weight: 640; }
    .stButton > button[kind="primary"] { background: var(--brand); border: none; }

    iframe { border: none !important; border-radius: 12px; }
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

logo_b64 = _logo_data_uri()
st.markdown(
    f"""
<div class="hdr">
    {f'<img src="data:image/png;base64,{logo_b64}"/>' if logo_b64 else ''}
    <div>
        <h1>Document Comparison Studio</h1>
        <p>Compare two PDF or Word files word by word. Every difference is highlighted in both.</p>
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Upload
# ---------------------------------------------------------------------------

st.markdown('<div class="step">Choose two documents <small>PDF or Word, in any combination</small></div>',
            unsafe_allow_html=True)

upload_left, upload_right = st.columns(2, gap="medium")

with upload_left:
    doc1_file = st.file_uploader(
        "Original", type=sorted(ALLOWED_EXTENSIONS), key="doc1"
    )
with upload_right:
    doc2_file = st.file_uploader(
        "Revised", type=sorted(ALLOWED_EXTENSIONS), key="doc2"
    )

ready = bool(doc1_file and doc2_file)
action_left, action_right, _ = st.columns([1.1, 1, 4])

with action_left:
    compare_clicked = st.button(
        "Compare documents", type="primary", width="stretch", disabled=not ready
    )
with action_right:
    if st.button("Clear", width="stretch"):
        st.session_state.pop("result", None)
        st.rerun()


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------

def send_mis_report(result: dict, bytes1: int, bytes2: int, started_at, duration: float) -> None:
    """Fire the MIS report. Back-office only: never surfaced in the interface."""
    try:
        headers = _safe_get_headers()
        ip_address = _extract_client_ip(headers)
        country, region, city, source = _resolve_location_from_ip(ip_address)
        location = ", ".join(p for p in [city, region, country] if p and p != "unknown") or "unknown"

        append_usage_event(
            {
                "event_id": str(uuid.uuid4()),
                "event_type": "comparison",
                "timestamp_utc": started_at.isoformat(),
                "session_id": _session_id(),
                "upload_count": 2,
                "comparison_mode": "streamlit",
                "doc1_name": result["doc1"]["name"],
                "doc2_name": result["doc2"]["name"],
                "doc1_type": result["doc1"]["extension"],
                "doc2_type": result["doc2"]["extension"],
                "duration_seconds": round(duration, 2),
                "highlighted_changes": result["summary"].get("highlighted_changes", 0),
                "match_rate": result["summary"].get("match_rate", 0),
                "client_ip": ip_address,
                "client_country": country,
                "client_region": region,
                "client_city": city,
                "location_source": source,
            }
        )

        if not mis_mailer.is_configured():
            logger.warning("MIS mail skipped: not configured")
            return

        sent, message = mis_mailer.send_mis_email(
            {
                "used_at": started_at,
                "duration_seconds": duration,
                "session_id": _session_id(),
                "client_ip": ip_address,
                "client_location": location,
                "environment": _environment_label(),
                "doc1": {
                    "name": result["doc1"]["name"],
                    "extension": result["doc1"]["extension"],
                    "size_bytes": bytes1,
                    "page_count": result["doc1"]["page_count"],
                    "output_name": result["doc1"]["output_name"],
                },
                "doc2": {
                    "name": result["doc2"]["name"],
                    "extension": result["doc2"]["extension"],
                    "size_bytes": bytes2,
                    "page_count": result["doc2"]["page_count"],
                    "output_name": result["doc2"]["output_name"],
                },
                "summary": result["summary"],
            }
        )
        (logger.info if sent else logger.warning)("MIS mail: %s", message)
    except Exception:
        # Reporting must never interfere with the comparison the user asked for.
        logger.exception("MIS reporting failed")


def run_comparison(file1, file2) -> None:
    bytes1 = file1.getvalue()
    bytes2 = file2.getvalue()
    started_at = datetime.now(timezone.utc)
    clock = time.perf_counter()

    progress = st.progress(0, text="Reading documents...")

    def on_progress(percent: int, label: str) -> None:
        progress.progress(min(percent, 100), text=label)

    try:
        result = compare_documents_full(
            file1.name, bytes1, file2.name, bytes2, progress_cb=on_progress
        )
    except ValueError as ex:
        progress.empty()
        st.error(str(ex), icon=":material/error:")
        return
    except Exception as ex:
        progress.empty()
        st.error(f"Comparison failed: {ex}", icon=":material/error:")
        return

    duration = time.perf_counter() - clock
    progress.progress(100, text="Finishing up...")

    result["meta"] = {"used_at": started_at, "duration_seconds": duration}
    send_mis_report(result, len(bytes1), len(bytes2), started_at, duration)

    progress.empty()
    st.session_state.result = result


if compare_clicked and ready:
    run_comparison(doc1_file, doc2_file)


# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------

result = st.session_state.get("result")

if result:
    summary = result["summary"]
    doc1, doc2 = result["doc1"], result["doc2"]

    largest = max(summary["total_words1"], summary["total_words2"])
    change_rate = (summary["highlighted_changes"] / largest * 100) if largest else 0

    st.markdown('<div class="step">What changed</div>', unsafe_allow_html=True)

    cards = [
        ("Words · original", format_indian_number(summary["total_words1"]), doc1["name"], ""),
        ("Words · revised", format_indian_number(summary["total_words2"]), doc2["name"], ""),
        ("Matching words", format_indian_number(summary["total_matching"]), "identical in both", ""),
        (
            "Highlighted changes",
            format_indian_number(summary["highlighted_changes"]),
            f"{format_indian_number(summary['diff_words1'])} original · "
            f"{format_indian_number(summary['diff_words2'])} revised",
            "warn",
        ),
        ("Match rate", f"{summary['match_rate']}%", f"change rate {change_rate:.1f}%", "ok"),
    ]
    st.markdown(
        '<div class="mrow">'
        + "".join(
            f'<div class="mcard {accent}"><div class="l">{label}</div>'
            f'<div class="v">{value}</div><div class="s">{sub}</div></div>'
            for label, value, sub, accent in cards
        )
        + "</div>",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="step">Download the highlighted files</div>', unsafe_allow_html=True)

    dl1, dl2, dl3, _ = st.columns([1.3, 1.3, 1.3, 2.1])
    with dl1:
        st.download_button(
            f"Original ({doc1['extension'].upper()})",
            data=doc1["output_bytes"], file_name=doc1["output_name"],
            mime=MIME_BY_EXT.get(doc1["extension"], "application/octet-stream"),
            width="stretch", key="dl_doc1",
        )
    with dl2:
        st.download_button(
            f"Revised ({doc2['extension'].upper()})",
            data=doc2["output_bytes"], file_name=doc2["output_name"],
            mime=MIME_BY_EXT.get(doc2["extension"], "application/octet-stream"),
            width="stretch", key="dl_doc2",
        )
    with dl3:
        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, "w", compression=ZIP_DEFLATED) as archive:
            archive.writestr(doc1["output_name"], doc1["output_bytes"])
            archive.writestr(doc2["output_name"], doc2["output_bytes"])
            archive.writestr("summary.json", json.dumps(summary, indent=2))
        st.download_button(
            "Both as ZIP", data=zip_buffer.getvalue(), file_name="comparison_result.zip",
            mime="application/zip", width="stretch", key="dl_zip",
        )

    st.markdown('<div class="step">Compare side by side '
                '<small>both panes scroll together</small></div>', unsafe_allow_html=True)

    st.markdown(
        '<div class="legend">'
        '<span><span class="sw" style="background:#fff59d"></span>Changed text</span>'
        '<span><span class="sw" style="background:#ffede0"></span>Changed inside a table</span>'
        '<span><span class="sw" style="background:#fff;border-color:#ea580c"></span>Current change</span>'
        '<span><span class="kbd">&larr;</span><span class="kbd">&rarr;</span> page'
        '&nbsp;&nbsp;<span class="kbd">N</span><span class="kbd">P</span> change'
        '&nbsp;&nbsp;<span class="kbd">S</span> sync</span>'
        "</div>",
        unsafe_allow_html=True,
    )

    st.iframe(viewer_data_url(build_viewer_html(doc1, doc2, summary)), height=VIEWER_HEIGHT)

elif ready:
    st.markdown(
        '<div class="empty">Both files are loaded &mdash; select <strong>Compare documents</strong>.</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<div class="empty">Upload an original and a revised document to begin. '
        "You can compare a PDF against a Word file.</div>",
        unsafe_allow_html=True,
    )
