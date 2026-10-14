"""Document Comparison Studio - Streamlit front-end.

Runs on the current comparison engine (``comparison_core``), the same one the
FastAPI/React build uses, so the Streamlit deployment and the API stay in step.
Each completed comparison is written to the usage log and e-mailed as an MIS
report (see ``mis_mailer``).
"""

from __future__ import annotations

import base64
import json
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
from usage_storage import get_usage_log_path

APP_DIR = Path(__file__).resolve().parent
USAGE_LOG_FILE = get_usage_log_path(__file__)
IST = timezone(timedelta(hours=5, minutes=30))

MIME_BY_EXT = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}

st.set_page_config(
    page_title="Document Comparison Studio",
    page_icon=str(APP_DIR / "logo.png") if (APP_DIR / "logo.png").exists() else ":material/difference:",
    layout="wide",
    initial_sidebar_state="expanded",
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


def format_bytes(size: int) -> str:
    if not size:
        return "-"
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.2f} MB"


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
    """Append to the JSONL the Usage Dashboard page reads."""
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
        --brand-dark: #004784;
        --brand-soft: #e8f2fb;
        --ink: #0f172a;
        --muted: #64748b;
        --line: #dde8f4;
        --panel: #ffffff;
        --warn: #b45309;
        --ok: #15803d;
    }

    html, body, [class*="css"] {
        font-family: "Inter", "Segoe UI", system-ui, Arial, sans-serif;
    }

    .stApp {
        background:
            radial-gradient(1100px 420px at 12% -5%, #e6f1fc 0%, rgba(230,241,252,0) 62%),
            linear-gradient(180deg, #f7fbff 0%, #ffffff 55%);
    }

    .block-container {
        padding-top: 1.6rem;
        padding-bottom: 3rem;
        max-width: 1500px;
    }

    #MainMenu, footer {visibility: hidden;}

    /* ---------- header ---------- */
    .brand-header {
        display: flex;
        align-items: center;
        gap: 20px;
        padding: 18px 24px;
        margin-bottom: 20px;
        border: 1px solid var(--line);
        border-radius: 18px;
        background: linear-gradient(135deg, #ffffff 0%, #f9fcff 100%);
        box-shadow: 0 10px 30px rgba(0, 94, 172, .07);
    }
    .brand-logo {
        width: 76px;
        height: 76px;
        object-fit: contain;
        flex-shrink: 0;
    }
    .brand-text h1 {
        color: var(--brand);
        margin: 0;
        font-size: 1.95rem;
        font-weight: 760;
        letter-spacing: -.015em;
        line-height: 1.1;
    }
    .brand-text p {
        color: var(--muted);
        margin: 5px 0 0 0;
        font-size: .94rem;
        line-height: 1.35;
    }
    .brand-pills {margin-left: auto; display: flex; gap: 8px; flex-wrap: wrap;}
    .pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        border-radius: 999px;
        padding: 6px 13px;
        font-size: .76rem;
        font-weight: 650;
        letter-spacing: .02em;
        border: 1px solid var(--line);
        background: var(--brand-soft);
        color: var(--brand-dark);
        white-space: nowrap;
    }
    .pill.ok {background:#ecfdf5; border-color:#a7f3d0; color: var(--ok);}
    .pill.warn {background:#fffbeb; border-color:#fde68a; color: var(--warn);}

    /* ---------- section heads ---------- */
    .section-head {
        display: flex;
        align-items: baseline;
        gap: 10px;
        margin: 26px 0 12px 0;
    }
    .section-head .num {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 24px;
        height: 24px;
        border-radius: 7px;
        background: var(--brand);
        color: #fff;
        font-size: .76rem;
        font-weight: 700;
        flex-shrink: 0;
        transform: translateY(2px);
    }
    .section-head h2 {
        margin: 0;
        font-size: 1.12rem;
        font-weight: 680;
        color: var(--ink);
        letter-spacing: -.01em;
    }
    .section-head span.hint {
        font-size: .83rem;
        color: var(--muted);
    }

    /* ---------- metric cards ---------- */
    .metric-row {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
        gap: 12px;
        margin: 4px 0 18px 0;
    }
    .metric-card {
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: 14px;
        padding: 15px 17px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, .045);
        border-top: 3px solid var(--brand);
    }
    .metric-card.accent-warn {border-top-color: #f59e0b;}
    .metric-card.accent-ok {border-top-color: #22c55e;}
    .metric-card .label {
        font-size: .73rem;
        text-transform: uppercase;
        letter-spacing: .08em;
        color: var(--muted);
        font-weight: 650;
    }
    .metric-card .value {
        font-size: 1.68rem;
        font-weight: 740;
        color: var(--ink);
        line-height: 1.2;
        margin-top: 4px;
        letter-spacing: -.02em;
    }
    .metric-card .sub {font-size: .76rem; color: var(--muted); margin-top: 2px;}
    .metric-card.accent-warn .value {color: var(--warn);}
    .metric-card.accent-ok .value {color: var(--ok);}

    /* ---------- file chip ---------- */
    .file-chip {
        display: flex;
        align-items: center;
        gap: 10px;
        border: 1px solid var(--line);
        background: #f8fbff;
        border-radius: 11px;
        padding: 9px 13px;
        margin-top: 8px;
    }
    .file-chip .ext {
        font-size: .68rem;
        font-weight: 750;
        letter-spacing: .05em;
        color: #fff;
        background: var(--brand);
        border-radius: 5px;
        padding: 3px 7px;
        flex-shrink: 0;
    }
    .file-chip .meta {overflow: hidden;}
    .file-chip .nm {
        font-size: .86rem;
        font-weight: 620;
        color: var(--ink);
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .file-chip .sz {font-size: .75rem; color: var(--muted);}

    /* ---------- legend ---------- */
    .legend {
        display: flex;
        gap: 18px;
        flex-wrap: wrap;
        align-items: center;
        border: 1px solid var(--line);
        background: var(--panel);
        border-radius: 12px;
        padding: 10px 15px;
        margin-bottom: 12px;
        font-size: .82rem;
        color: var(--muted);
    }
    .swatch {
        display: inline-block;
        width: 15px;
        height: 15px;
        border-radius: 4px;
        vertical-align: -3px;
        margin-right: 7px;
        border: 1px solid rgba(0,0,0,.12);
    }

    /* ---------- preview ---------- */
    .doc-head {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        border: 1px solid var(--line);
        border-bottom: none;
        border-radius: 12px 12px 0 0;
        background: linear-gradient(180deg, #ffffff, #f7fbff);
        padding: 11px 15px;
    }
    .doc-head .nm {
        font-size: .9rem;
        font-weight: 680;
        color: var(--ink);
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .doc-head .ct {
        font-size: .76rem;
        color: var(--muted);
        white-space: nowrap;
        flex-shrink: 0;
    }

    /* ---------- streamlit widget polish ---------- */
    div[data-testid="stFileUploader"] section {
        border: 1.5px dashed #a9cdec !important;
        border-radius: 13px !important;
        background: #fbfdff !important;
        padding: 14px !important;
        transition: border-color .15s ease, background .15s ease;
    }
    div[data-testid="stFileUploader"] section:hover {
        border-color: var(--brand) !important;
        background: #f5faff !important;
    }
    div[data-testid="stFileUploader"] small {color: var(--muted);}

    .stButton > button, .stDownloadButton > button {
        border-radius: 11px;
        font-weight: 640;
        border: 1px solid var(--line);
        transition: transform .08s ease, box-shadow .15s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(0, 94, 172, .16);
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--brand) 0%, #0b80d8 100%);
        border: none;
    }

    div[data-testid="stSidebarContent"] {
        background: #f7fbff;
        border-right: 1px solid var(--line);
    }

    div[data-testid="stImageContainer"] img, .pdf-page img {
        border-radius: 8px;
        box-shadow: 0 5px 16px rgba(15, 23, 42, .11);
        border: 1px solid var(--line);
    }

    hr {border-color: var(--line);}
</style>
""",
    unsafe_allow_html=True,
)


def section_head(number: str, title: str, hint: str = "") -> None:
    st.markdown(
        f'<div class="section-head"><span class="num">{number}</span>'
        f'<h2>{title}</h2><span class="hint">{hint}</span></div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

logo_b64 = _logo_data_uri()
mis_ready = mis_mailer.is_configured()
mis_pill = (
    '<span class="pill ok">&#9679; MIS reporting on</span>'
    if mis_ready
    else '<span class="pill warn">&#9679; MIS not configured</span>'
)
logo_html = f'<img src="data:image/png;base64,{logo_b64}" class="brand-logo"/>' if logo_b64 else ""

st.markdown(
    f"""
<div class="brand-header">
    {logo_html}
    <div class="brand-text">
        <h1>Document Comparison Studio</h1>
        <p>Compare PDF and DOCX files word by word, with every difference highlighted in place.</p>
    </div>
    <div class="brand-pills">
        <span class="pill">PDF &middot; DOCX</span>
        <span class="pill">Mixed formats</span>
        {mis_pill}
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown("### :material/tune: Settings")

    view_mode = st.radio(
        "Preview layout",
        ["Side by side", "Stacked"],
        help="Side by side puts both documents in adjacent columns. Stacked gives each "
             "document the full width - easier on a laptop screen.",
    )

    preview_height = st.slider("Preview height (px)", 420, 1200, 760, step=20)

    max_pages = st.number_input(
        "Max PDF pages to render",
        min_value=0,
        max_value=500,
        value=0,
        step=5,
        help="0 renders every page. Lower this for very long PDFs to speed up the preview. "
             "The downloadable highlighted files always contain all pages.",
    )

    st.divider()
    st.markdown("### :material/mail: MIS reporting")
    send_mis = st.toggle("E-mail an MIS report per comparison", value=True, disabled=not mis_ready)

    if mis_ready:
        st.caption("Recipients")
        for address in mis_mailer.recipients():
            st.caption(f"&nbsp;&nbsp;:material/person: {address}", unsafe_allow_html=True)
    else:
        st.warning(
            "Add `GMAIL_USERNAME`, `GMAIL_APP_PASSWORD` and `MIS_EMAIL_TO` to "
            "`.streamlit/secrets.toml` to enable MIS mails.",
            icon=":material/key_off:",
        )

    st.divider()
    st.caption(
        "Highlights: **yellow** for changed body text, **peach** for changes inside tables. "
        "Isolated one-off differences are filtered out to cut formatting noise."
    )


# ---------------------------------------------------------------------------
# Upload
# ---------------------------------------------------------------------------

section_head("1", "Upload documents", "PDF or DOCX, in any combination")

upload_left, upload_right = st.columns(2, gap="large")


def file_chip(uploaded) -> None:
    if not uploaded:
        return
    ext = Path(uploaded.name).suffix.lstrip(".").upper() or "FILE"
    st.markdown(
        f'<div class="file-chip"><span class="ext">{ext}</span>'
        f'<div class="meta"><div class="nm">{uploaded.name}</div>'
        f'<div class="sz">{format_bytes(uploaded.size)}</div></div></div>',
        unsafe_allow_html=True,
    )


with upload_left:
    st.markdown("**:material/description: Document 1** &nbsp;·&nbsp; original / baseline", unsafe_allow_html=True)
    doc1_file = st.file_uploader(
        "Document 1", type=sorted(ALLOWED_EXTENSIONS), key="doc1", label_visibility="collapsed"
    )
    file_chip(doc1_file)

with upload_right:
    st.markdown("**:material/description: Document 2** &nbsp;·&nbsp; revised / new version", unsafe_allow_html=True)
    doc2_file = st.file_uploader(
        "Document 2", type=sorted(ALLOWED_EXTENSIONS), key="doc2", label_visibility="collapsed"
    )
    file_chip(doc2_file)

action_left, action_right, _ = st.columns([1.1, 1, 3])
ready = bool(doc1_file and doc2_file)

with action_left:
    compare_clicked = st.button(
        ":material/compare_arrows: Compare documents",
        type="primary",
        width="stretch",
        disabled=not ready,
    )

with action_right:
    if st.button(":material/restart_alt: Clear results", width="stretch"):
        st.session_state.pop("result", None)
        st.session_state.pop("mis_status", None)
        st.rerun()

if not ready:
    st.info("Upload both documents to start a comparison.", icon=":material/upload_file:")


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------

def run_comparison(file1, file2, page_limit: int, email_report: bool) -> None:
    bytes1 = file1.getvalue()
    bytes2 = file2.getvalue()
    started_at = datetime.now(timezone.utc)
    clock = time.perf_counter()

    progress = st.progress(0, text="Starting comparison...")

    def on_progress(percent: int, label: str) -> None:
        progress.progress(min(percent, 100), text=label)

    try:
        result = compare_documents_full(
            file1.name, bytes1, file2.name, bytes2, max_pages=page_limit, progress_cb=on_progress
        )
    except ValueError as ex:
        progress.empty()
        st.error(str(ex), icon=":material/error:")
        return
    except Exception as ex:
        progress.empty()
        st.error(f"Comparison failed: {ex}", icon=":material/error:")
        return

    progress.empty()
    duration = time.perf_counter() - clock

    result["doc1"]["size_bytes"] = len(bytes1)
    result["doc2"]["size_bytes"] = len(bytes2)
    result["meta"] = {"used_at": started_at, "duration_seconds": duration}

    headers = _safe_get_headers()
    ip_address = _extract_client_ip(headers)
    country, region, city, source = _resolve_location_from_ip(ip_address)
    location = ", ".join(part for part in [city, region, country] if part and part != "unknown") or "unknown"

    summary = result["summary"]

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
            "highlighted_changes": summary.get("highlighted_changes", 0),
            "match_rate": summary.get("match_rate", 0),
            "client_ip": ip_address,
            "client_country": country,
            "client_region": region,
            "client_city": city,
            "location_source": source,
        }
    )

    st.session_state.result = result
    st.session_state.mis_status = None

    if email_report:
        with st.spinner("Sending MIS report..."):
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
                        "size_bytes": len(bytes1),
                        "page_count": result["doc1"]["page_count"],
                        "output_name": result["doc1"]["output_name"],
                    },
                    "doc2": {
                        "name": result["doc2"]["name"],
                        "extension": result["doc2"]["extension"],
                        "size_bytes": len(bytes2),
                        "page_count": result["doc2"]["page_count"],
                        "output_name": result["doc2"]["output_name"],
                    },
                    "summary": summary,
                }
            )
        st.session_state.mis_status = (sent, message)


if compare_clicked and ready:
    run_comparison(doc1_file, doc2_file, int(max_pages), send_mis and mis_ready)


# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------

def metric_cards(cards: list[tuple[str, str, str, str]]) -> None:
    """cards = [(label, value, sub, accent_class)]"""
    html = "".join(
        f'<div class="metric-card {accent}"><div class="label">{label}</div>'
        f'<div class="value">{value}</div><div class="sub">{sub}</div></div>'
        for label, value, sub, accent in cards
    )
    st.markdown(f'<div class="metric-row">{html}</div>', unsafe_allow_html=True)


def render_preview(doc: dict, height: int) -> None:
    preview = doc.get("preview") or {}
    pages = preview.get("pages") or []
    total = preview.get("total_pages") or 0

    count_label = (
        f"{format_indian_number(len(pages))} of {format_indian_number(total)} pages"
        if preview.get("type") == "pdf_images"
        else "Rendered from DOCX"
    )

    st.markdown(
        f'<div class="doc-head"><span class="nm">{doc["name"]}</span>'
        f'<span class="ct">{count_label}</span></div>',
        unsafe_allow_html=True,
    )

    if preview.get("type") == "html":
        # st.html inside a fixed-height container scrolls natively and picks up
        # the page theme, unlike an iframe which needs its height guessed.
        with st.container(height=height, border=True):
            st.html(preview.get("html", ""))
        return

    with st.container(height=height, border=True):
        if not pages:
            st.info("No pages to preview.")
            return
        for page in pages:
            encoded = page.get("image_base64")
            if not encoded:
                continue
            st.image(base64.b64decode(encoded), width="stretch")
            st.caption(f"Page {format_indian_number(page['page'])} of {format_indian_number(total)}")

    if preview.get("truncated"):
        st.caption(
            f":material/info: Preview limited to {format_indian_number(len(pages))} pages. "
            "The downloadable file still contains every page."
        )


result = st.session_state.get("result")

if result:
    summary = result["summary"]
    doc1, doc2 = result["doc1"], result["doc2"]
    meta = result["meta"]

    used_ist = meta["used_at"].astimezone(IST)
    st.success(
        f"Comparison complete in {meta['duration_seconds']:.1f}s "
        f"· {used_ist.strftime('%d %b %Y, %I:%M %p')} IST",
        icon=":material/task_alt:",
    )

    mis_status = st.session_state.get("mis_status")
    if mis_status:
        sent, message = mis_status
        if sent:
            st.caption(f":material/mark_email_read: {message}")
        else:
            st.warning(message, icon=":material/unsubscribe:")

    section_head("2", "Comparison summary")

    largest = max(summary["total_words1"], summary["total_words2"])
    change_rate = (summary["highlighted_changes"] / largest * 100) if largest else 0

    metric_cards(
        [
            ("Words · doc 1", format_indian_number(summary["total_words1"]), doc1["name"][:34], ""),
            ("Words · doc 2", format_indian_number(summary["total_words2"]), doc2["name"][:34], ""),
            ("Matching words", format_indian_number(summary["total_matching"]), "identical in both", ""),
            (
                "Highlighted changes",
                format_indian_number(summary["highlighted_changes"]),
                f"{format_indian_number(summary['diff_words1'])} in doc 1 · "
                f"{format_indian_number(summary['diff_words2'])} in doc 2",
                "accent-warn",
            ),
            ("Match rate", f"{summary['match_rate']}%", f"change rate {change_rate:.1f}%", "accent-ok"),
        ]
    )

    st.progress(min(summary["match_rate"] / 100, 1.0), text=f"Documents match {summary['match_rate']}%")

    section_head("3", "Download highlighted documents")

    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        st.download_button(
            f":material/file_download: {doc1['extension'].upper()} · document 1",
            data=doc1["output_bytes"],
            file_name=doc1["output_name"],
            mime=MIME_BY_EXT.get(doc1["extension"], "application/octet-stream"),
            width="stretch",
            key="dl_doc1",
        )

    with dl2:
        st.download_button(
            f":material/file_download: {doc2['extension'].upper()} · document 2",
            data=doc2["output_bytes"],
            file_name=doc2["output_name"],
            mime=MIME_BY_EXT.get(doc2["extension"], "application/octet-stream"),
            width="stretch",
            key="dl_doc2",
        )

    with dl3:
        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, "w", compression=ZIP_DEFLATED) as archive:
            archive.writestr(doc1["output_name"], doc1["output_bytes"])
            archive.writestr(doc2["output_name"], doc2["output_bytes"])
            archive.writestr("summary.json", json.dumps(summary, indent=2))
        st.download_button(
            ":material/folder_zip: Both files + summary (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="comparison_result.zip",
            mime="application/zip",
            width="stretch",
            key="dl_zip",
        )

    section_head("4", "Highlighted preview")

    st.markdown(
        '<div class="legend">'
        '<span><span class="swatch" style="background:#fff59d"></span>Changed text</span>'
        '<span><span class="swatch" style="background:#ffede0"></span>Changed text inside a table</span>'
        '<span><span class="swatch" style="background:#ffffff"></span>Unchanged</span>'
        "</div>",
        unsafe_allow_html=True,
    )

    if view_mode == "Side by side":
        preview_left, preview_right = st.columns(2, gap="medium")
        with preview_left:
            render_preview(doc1, preview_height)
        with preview_right:
            render_preview(doc2, preview_height)
    else:
        render_preview(doc1, preview_height)
        st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
        render_preview(doc2, preview_height)

elif ready:
    st.markdown(
        '<div style="border:1px solid #dde8f4;border-radius:14px;padding:26px;background:#fff;'
        'text-align:center;color:#64748b;margin-top:14px">'
        "Both documents are ready. Select <strong>Compare documents</strong> to run the comparison."
        "</div>",
        unsafe_allow_html=True,
    )
