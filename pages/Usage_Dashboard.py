import json
import os
from pathlib import Path

import streamlit as st

st.set_page_config(page_title="Usage Dashboard", layout="wide")
st.title("📊 Usage Dashboard")
st.caption("Admin-only analytics for document comparison usage")


def _get_dashboard_password():
    try:
        secret_password = st.secrets.get("ADMIN_DASHBOARD_PASSWORD")
        if secret_password:
            return str(secret_password)
    except Exception:
        pass

    env_password = os.getenv("ADMIN_DASHBOARD_PASSWORD")
    return str(env_password) if env_password else None


def _usage_log_path():
    project_root = Path(__file__).resolve().parent.parent
    return project_root / "usage_logs.jsonl"


def load_usage_logs():
    log_file = _usage_log_path()
    if not log_file.exists():
        return []

    records = []
    try:
        with open(log_file, "r", encoding="utf-8") as file:
            for line in file:
                line = line.strip()
                if not line:
                    continue
                try:
                    records.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
    except OSError:
        return []

    records.sort(key=lambda item: item.get("timestamp_utc", ""), reverse=True)
    return records


def render_usage_dashboard():
    logs = load_usage_logs()
    if not logs:
        st.info("No usage data yet. Upload and compare documents to start tracking usage.")
        return

    total_events = len(logs)
    total_uploads = sum(int(item.get("upload_count", 0)) for item in logs)
    unique_ips = len(
        {
            item.get("client_ip")
            for item in logs
            if item.get("client_ip") and item.get("client_ip") != "unknown"
        }
    )
    last_seen = logs[0].get("timestamp_utc", "-")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Comparisons", total_events)
    m2.metric("Total Uploads", total_uploads)
    m3.metric("Unique IPs", unique_ips)
    m4.metric("Last Activity (UTC)", last_seen.replace("T", " ")[:19] if last_seen != "-" else "-")

    daily_uploads = {}
    for item in logs:
        date_key = item.get("timestamp_utc", "")[:10]
        if not date_key:
            continue
        daily_uploads[date_key] = daily_uploads.get(date_key, 0) + int(item.get("upload_count", 0))

    if daily_uploads:
        recent_days = dict(sorted(daily_uploads.items())[-30:])
        st.markdown("**Uploads per day (last 30 days)**")
        st.bar_chart(recent_days)

    location_counts = {}
    for item in logs:
        city = item.get("client_city", "unknown")
        country = item.get("client_country", "unknown")
        label = f"{city}, {country}"
        location_counts[label] = location_counts.get(label, 0) + 1

    top_locations = sorted(location_counts.items(), key=lambda x: x[1], reverse=True)[:10]
    st.markdown("**Top locations (best effort)**")
    st.dataframe(
        [{"location": name, "comparisons": count} for name, count in top_locations],
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("**Recent usage events**")
    st.dataframe(
        [
            {
                "timestamp_utc": item.get("timestamp_utc", ""),
                "doc1": item.get("doc1_name", ""),
                "doc2": item.get("doc2_name", ""),
                "mode": item.get("comparison_mode", ""),
                "ip": item.get("client_ip", "unknown"),
                "location": f"{item.get('client_city', 'unknown')}, {item.get('client_country', 'unknown')}",
            }
            for item in logs[:100]
        ],
        use_container_width=True,
        hide_index=True,
    )


configured_password = _get_dashboard_password()
if not configured_password:
    st.warning("Dashboard is disabled. Set ADMIN_DASHBOARD_PASSWORD in Streamlit secrets or environment variables.")
    st.stop()

if "dashboard_authenticated" not in st.session_state:
    st.session_state.dashboard_authenticated = False

with st.sidebar:
    st.markdown("### Admin Access")
    password_input = st.text_input("Dashboard password", type="password", key="dashboard_password_input")
    login_col, logout_col = st.columns(2)

    with login_col:
        if st.button("Unlock", key="dashboard_unlock"):
            st.session_state.dashboard_authenticated = password_input == configured_password
            if not st.session_state.dashboard_authenticated:
                st.error("Invalid password")

    with logout_col:
        if st.button("Lock", key="dashboard_lock"):
            st.session_state.dashboard_authenticated = False

if not st.session_state.dashboard_authenticated:
    st.info("Enter admin password in the sidebar to view usage analytics.")
    st.stop()

render_usage_dashboard()
