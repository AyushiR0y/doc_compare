"""Document Comparison Studio - Streamlit front-end.

Runs on the current comparison engine (``comparison_core``), the same one the
FastAPI/React build uses, so the Streamlit deployment and the API stay in step.
Each completed comparison is written to the usage log and e-mailed as an MIS
report (see ``mis_mailer``).

The side-by-side preview lives in ``preview_viewer`` - Streamlit cannot
synchronise scrolling between two of its own containers, so both documents and
their controls are rendered into one embedded HTML document.
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
from preview_viewer import build_viewer_html, viewer_data_url
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
        --brand-dark: #00457f;
        --brand-soft: #e9f3fc;
        --ink: #0f172a;
        --muted: #64748b;
        --faint: #94a3b8;
        --line: #dde8f4;
        --panel: #ffffff;
        --warn: #b45309;
        --ok: #15803d;
        --radius: 14px;
    }

    html, body, [class*="css"] { font-family: "Inter", "Segoe UI", system-ui, Arial, sans-serif; }

    .stApp {
        background:
            radial-gradient(900px 360px at 8% -8%, #e4effb 0%, rgba(228,239,251,0) 60%),
            radial-gradient(700px 320px at 98% 2%, #eaf4ec 0%, rgba(234,244,236,0) 55%),
            linear-gradient(180deg, #f8fbff 0%, #ffffff 42%);
    }

    /* the default top padding clipped the header card */
    .block-container { padding: 2.4rem 2.2rem 3rem; max-width: 1680px; }
    header[data-testid="stHeader"] { background: transparent; }
    #MainMenu, footer { visibility: hidden; }

    /* ---------- hero ---------- */
    .hero {
        display: flex; align-items: center; gap: 18px;
        padding: 16px 22px; margin-bottom: 4px;
        border: 1px solid var(--line); border-radius: 18px;
        background: linear-gradient(120deg, #ffffff 0%, #f7fbff 55%, #f2f8ff 100%);
        box-shadow: 0 10px 28px rgba(0, 94, 172, .08);
    }
    .hero-logo { width: 58px; height: 58px; object-fit: contain; flex: 0 0 auto; }
    .hero-txt h1 {
        color: var(--brand); margin: 0; font-size: 1.62rem; font-weight: 750;
        letter-spacing: -.02em; line-height: 1.12;
    }
    .hero-txt p { color: var(--muted); margin: 4px 0 0; font-size: .89rem; line-height: 1.3; }
    .hero-pills { margin-left: auto; display: flex; gap: 7px; flex-wrap: wrap; justify-content: flex-end; }
    .pill {
        display: inline-flex; align-items: center; border-radius: 999px; padding: 5px 12px;
        font-size: .73rem; font-weight: 680; letter-spacing: .015em; white-space: nowrap;
        border: 1px solid var(--line); background: var(--brand-soft); color: var(--brand-dark);
    }
    .pill.ok { background: #ecfdf5; border-color: #a7f3d0; color: var(--ok); }
    .pill.warn { background: #fffbeb; border-color: #fde68a; color: var(--warn); }

    /* ---------- section heads ---------- */
    .shead { display: flex; align-items: baseline; gap: 10px; margin: 22px 0 10px; }
    .shead .n {
        display: inline-flex; align-items: center; justify-content: center;
        width: 22px; height: 22px; border-radius: 7px; background: var(--brand);
        color: #fff; font-size: .72rem; font-weight: 750; flex: 0 0 auto; transform: translateY(2px);
    }
    .shead h2 { margin: 0; font-size: 1.05rem; font-weight: 700; color: var(--ink); letter-spacing: -.01em; }
    .shead .hint { font-size: .81rem; color: var(--faint); }

    /* ---------- upload cards ---------- */
    .ucard-head { display: flex; align-items: baseline; gap: 8px; margin-bottom: 4px; }
    .ucard-dot {
        width: 22px; height: 22px; border-radius: 7px; background: var(--brand-soft);
        color: var(--brand); display: inline-flex; align-items: center; justify-content: center;
        font-size: .72rem; font-weight: 800; flex: 0 0 auto; transform: translateY(2px);
    }
    .ucard-head b { font-size: .92rem; color: var(--ink); font-weight: 700; }
    .ucard-head span { font-size: .78rem; color: var(--faint); }

    /* ---------- metric cards ---------- */
    .mrow { display: grid; grid-template-columns: repeat(auto-fit, minmax(158px, 1fr)); gap: 11px; margin-bottom: 14px; }
    .mcard {
        background: var(--panel); border: 1px solid var(--line); border-radius: var(--radius);
        padding: 13px 15px; box-shadow: 0 3px 12px rgba(15, 23, 42, .045); position: relative;
        overflow: hidden;
    }
    .mcard::before {
        content: ""; position: absolute; inset: 0 0 auto 0; height: 3px; background: var(--brand);
    }
    .mcard.warn::before { background: linear-gradient(90deg, #f59e0b, #fbbf24); }
    .mcard.ok::before { background: linear-gradient(90deg, #22c55e, #4ade80); }
    .mcard .l {
        font-size: .69rem; text-transform: uppercase; letter-spacing: .08em;
        color: var(--faint); font-weight: 700;
    }
    .mcard .v {
        font-size: 1.55rem; font-weight: 760; color: var(--ink); line-height: 1.18;
        margin-top: 3px; letter-spacing: -.025em;
    }
    .mcard .s { font-size: .73rem; color: var(--muted); margin-top: 1px;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .mcard.warn .v { color: var(--warn); }
    .mcard.ok .v { color: var(--ok); }

    .meter { height: 7px; border-radius: 999px; background: #eef2f7; overflow: hidden; margin: 2px 0 16px; }
    .meter > i { display: block; height: 100%; border-radius: 999px;
        background: linear-gradient(90deg, var(--brand), #3fa0e8); }

    /* ---------- legend + empty ---------- */
    .legend {
        display: flex; gap: 16px; flex-wrap: wrap; align-items: center;
        border: 1px solid var(--line); background: var(--panel); border-radius: 11px;
        padding: 9px 14px; margin-bottom: 10px; font-size: .79rem; color: var(--muted);
    }
    .sw { display: inline-block; width: 13px; height: 13px; border-radius: 4px;
        vertical-align: -2px; margin-right: 6px; border: 1px solid rgba(0,0,0,.12); }
    .kbd {
        font-family: ui-monospace, Menlo, Consolas, monospace; font-size: .7rem;
        border: 1px solid var(--line); border-bottom-width: 2px; border-radius: 5px;
        padding: 1px 5px; background: #f8fafc; color: var(--muted); margin: 0 1px;
    }
    .empty {
        border: 1.5px dashed #c6dcf0; border-radius: var(--radius); background: #fbfdff;
        padding: 34px 24px; text-align: center; color: var(--muted); margin-top: 10px;
    }
    .empty b { display: block; color: var(--ink); font-size: 1rem; margin-bottom: 4px; font-weight: 700; }
    .empty span { font-size: .86rem; }

    /* ---------- widget polish ---------- */
    div[data-testid="stFileUploader"] section {
        border: 1.5px dashed #a9cdec !important; border-radius: 11px !important;
        background: #fbfdff !important; padding: 11px 13px !important;
        transition: border-color .14s ease, background .14s ease;
    }
    div[data-testid="stFileUploader"] section:hover { border-color: var(--brand) !important; background: #f4faff !important; }
    div[data-testid="stFileUploader"] small { color: var(--faint); }

    .stButton > button, .stDownloadButton > button {
        border-radius: 10px; font-weight: 640; border: 1px solid var(--line);
        transition: transform .08s ease, box-shadow .14s ease, background .14s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-1px); box-shadow: 0 6px 15px rgba(0, 94, 172, .15);
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--brand) 0%, #0b80d8 100%); border: none;
    }

    div[data-testid="stSidebarContent"] { background: #f7fbff; border-right: 1px solid var(--line); }
    div[data-testid="stSidebarContent"] hr { margin: .9rem 0; }

    iframe { border: none !important; border-radius: var(--radius); }
</style>
""",
    unsafe_allow_html=True,
)


def shead(number: str, title: str, hint: str = "") -> None:
    st.markdown(
        f'<div class="shead"><span class="n">{number}</span><h2>{title}</h2>'
        f'<span class="hint">{hint}</span></div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Sidebar
#
# Rendered before the hero so the hero's MIS pill can show whether reporting is
# actually active, not merely configured. Streamlit places this in the sidebar
# regardless of where it runs in the script.
# ---------------------------------------------------------------------------

mis_ready = mis_mailer.is_configured()

with st.sidebar:
    st.page_link("streamlit_app.py", label="Compare documents", icon=":material/difference:")
    st.page_link("pages/Usage_Dashboard.py", label="Usage dashboard", icon=":material/monitoring:")
    st.divider()

    st.markdown("##### :material/tune: Viewer")
    layout_mode = st.radio(
        "Layout",
        ["Side by side", "Stacked"],
        horizontal=True,
        help="Side by side is best on a wide screen. Stacked gives each document the full width.",
    )
    viewer_height = st.slider("Viewer height (px)", 520, 1400, 860, step=20)
    max_pages = st.number_input(
        "Max PDF pages to render",
        min_value=0, max_value=500, value=0, step=5,
        help="0 renders every page. Lower this for very long PDFs to speed up the preview. "
             "The downloadable highlighted files always contain all pages.",
    )

    st.divider()
    st.markdown("##### :material/mail: MIS reporting")
    send_mis = st.toggle("E-mail a report per comparison", value=True, disabled=not mis_ready)
    if mis_ready:
        st.caption("Recipients: " + ", ".join(mis_mailer.recipients()))
    else:
        st.warning(
            "Add `GMAIL_APP_PASSWORD` to `.streamlit/secrets.toml` to enable MIS mails.",
            icon=":material/key_off:",
        )

    st.divider()
    st.markdown("##### :material/keyboard: Viewer shortcuts")
    st.markdown(
        '<div style="font-size:.79rem;line-height:2.1;color:#64748b">'
        '<span class="kbd">&larr;</span><span class="kbd">&rarr;</span> previous / next page<br>'
        '<span class="kbd">N</span><span class="kbd">P</span> next / previous change<br>'
        '<span class="kbd">S</span> toggle synchronised scrolling'
        "</div>",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Hero
# ---------------------------------------------------------------------------

logo_b64 = _logo_data_uri()
if not mis_ready:
    mis_pill = '<span class="pill warn">&#9679;&nbsp; MIS not configured</span>'
elif send_mis:
    mis_pill = '<span class="pill ok">&#9679;&nbsp; MIS reporting on</span>'
else:
    mis_pill = '<span class="pill">&#9675;&nbsp; MIS reporting paused</span>'

st.markdown(
    f"""
<div class="hero">
    {f'<img src="data:image/png;base64,{logo_b64}" class="hero-logo"/>' if logo_b64 else ''}
    <div class="hero-txt">
        <h1>Document Comparison Studio</h1>
        <p>Compare PDF and DOCX files word by word, with every difference highlighted in place.</p>
    </div>
    <div class="hero-pills">
        <span class="pill">PDF &middot; DOCX &middot; mixed</span>
        {mis_pill}
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Upload
# ---------------------------------------------------------------------------

shead("1", "Upload documents", "PDF or DOCX, in any combination")

upload_left, upload_right = st.columns(2, gap="large")


def upload_card(column, number: str, title: str, caption: str, key: str):
    with column:
        st.markdown(
            f'<div class="ucard-head"><span class="ucard-dot">{number}</span>'
            f'<b>{title}</b><span>&middot; {caption}</span></div>',
            unsafe_allow_html=True,
        )
        # The uploader already lists the chosen file with its size and a remove
        # control, so there is no second chip here repeating it.
        return st.file_uploader(
            title, type=sorted(ALLOWED_EXTENSIONS), key=key, label_visibility="collapsed"
        )


doc1_file = upload_card(upload_left, "1", "Document 1", "original / baseline", "doc1")
doc2_file = upload_card(upload_right, "2", "Document 2", "revised / new version", "doc2")

action_left, action_right, _ = st.columns([1.15, 1, 3.2])
ready = bool(doc1_file and doc2_file)

with action_left:
    compare_clicked = st.button(
        ":material/compare_arrows: Compare documents",
        type="primary", width="stretch", disabled=not ready,
    )

with action_right:
    if st.button(":material/restart_alt: Clear results", width="stretch"):
        st.session_state.pop("result", None)
        st.session_state.pop("mis_status", None)
        st.rerun()


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
    location = ", ".join(p for p in [city, region, country] if p and p != "unknown") or "unknown"

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
    html = "".join(
        f'<div class="mcard {accent}"><div class="l">{label}</div>'
        f'<div class="v">{value}</div><div class="s">{sub}</div></div>'
        for label, value, sub, accent in cards
    )
    st.markdown(f'<div class="mrow">{html}</div>', unsafe_allow_html=True)


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

    shead("2", "Summary")

    largest = max(summary["total_words1"], summary["total_words2"])
    change_rate = (summary["highlighted_changes"] / largest * 100) if largest else 0

    metric_cards(
        [
            ("Words · doc 1", format_indian_number(summary["total_words1"]), doc1["name"], ""),
            ("Words · doc 2", format_indian_number(summary["total_words2"]), doc2["name"], ""),
            ("Matching words", format_indian_number(summary["total_matching"]), "identical in both", ""),
            (
                "Highlighted changes",
                format_indian_number(summary["highlighted_changes"]),
                f"{format_indian_number(summary['diff_words1'])} in doc 1 · "
                f"{format_indian_number(summary['diff_words2'])} in doc 2",
                "warn",
            ),
            ("Match rate", f"{summary['match_rate']}%", f"change rate {change_rate:.1f}%", "ok"),
        ]
    )
    st.markdown(
        f'<div class="meter"><i style="width:{min(summary["match_rate"], 100):.1f}%"></i></div>',
        unsafe_allow_html=True,
    )

    shead("3", "Download highlighted documents")

    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.download_button(
            f":material/file_download: {doc1['extension'].upper()} · document 1",
            data=doc1["output_bytes"], file_name=doc1["output_name"],
            mime=MIME_BY_EXT.get(doc1["extension"], "application/octet-stream"),
            width="stretch", key="dl_doc1",
        )
    with dl2:
        st.download_button(
            f":material/file_download: {doc2['extension'].upper()} · document 2",
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
            ":material/folder_zip: Both files + summary (ZIP)",
            data=zip_buffer.getvalue(), file_name="comparison_result.zip",
            mime="application/zip", width="stretch", key="dl_zip",
        )

    shead("4", "Side-by-side preview", "scrolling stays in sync · step through changes")

    st.markdown(
        '<div class="legend">'
        '<span><span class="sw" style="background:#fff59d"></span>Changed text</span>'
        '<span><span class="sw" style="background:#ffede0"></span>Changed text inside a table</span>'
        '<span><span class="sw" style="background:#fff;border-color:#ea580c"></span>Current change</span>'
        '<span style="margin-left:auto">'
        '<span class="kbd">&larr;</span><span class="kbd">&rarr;</span> page'
        '&nbsp;&nbsp;<span class="kbd">N</span><span class="kbd">P</span> change'
        '&nbsp;&nbsp;<span class="kbd">S</span> sync</span>'
        "</div>",
        unsafe_allow_html=True,
    )

    viewer_html = build_viewer_html(doc1, doc2, summary, stacked=(layout_mode == "Stacked"))
    st.iframe(viewer_data_url(viewer_html), height=viewer_height)

    truncated = [d["name"] for d in (doc1, doc2) if (d.get("preview") or {}).get("truncated")]
    if truncated:
        st.caption(
            f":material/info: Preview limited by the page cap for {', '.join(truncated)}. "
            "The downloadable files still contain every page."
        )

elif ready:
    st.markdown(
        '<div class="empty"><b>Ready to compare</b>'
        "<span>Both documents are loaded. Select <strong>Compare documents</strong> to run it.</span></div>",
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<div class="empty"><b>Upload two documents to begin</b>'
        "<span>PDF or DOCX, in any combination &mdash; you can compare a PDF against a Word file.</span></div>",
        unsafe_allow_html=True,
    )
