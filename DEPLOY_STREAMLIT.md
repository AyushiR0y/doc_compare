# Deploying Document Comparison Studio to Streamlit

## Which file is the app?

| File | What it is | Status |
| --- | --- | --- |
| `streamlit_app.py` | **Deploy this.** Streamlit UI on the current engine, with MIS e-mail. | Current |
| `comparison_core.py` | The comparison engine. Shared by the Streamlit app and the API. | Current |
| `mis_mailer.py` | Builds and sends the MIS report. | Current |
| `pages/Usage_Dashboard.py` | Usage analytics page, appears in the sidebar automatically. | Current |
| `fastapi_app.py` + `frontend/` | FastAPI + React build of the same engine (Render deployment). | Parallel deployment |
| `app.py` | The previous Streamlit app, with its own older copy of the comparison logic. | Superseded |

`app.py` is left in place but is no longer the entry point. It carries a duplicate,
older version of the diff logic, so it will not pick up engine improvements.

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy to Streamlit Community Cloud

1. Push the repository to GitHub (`.streamlit/secrets.toml` is git-ignored and must
   stay that way).
2. On <https://share.streamlit.io> choose **New app**, pick the repo and branch.
3. Set **Main file path** to `doc_compare/streamlit_app.py`.
4. Open **Advanced settings -> Python version** and select **3.10** or later.
5. Open **Settings -> Secrets** and paste:

   ```toml
   GMAIL_USERNAME = "commissionsmis@gmail.com"
   GMAIL_APP_PASSWORD = "your-16-character-app-password"
   MIS_EMAIL_TO = "Ayushi.Roy01@bajajlife.com,bhagyashree.daga@bajajlife.com"
   ```

6. Deploy. The header shows a green **MIS reporting on** pill once the secrets
   resolve; if it shows an amber **MIS not configured** pill, the secrets are missing.

`requirements.txt` is shared with the FastAPI deployment, so it also installs
`fastapi`/`uvicorn`. They are unused by the Streamlit app and harmless.

## MIS reporting

Every completed comparison sends one mail to every address in `MIS_EMAIL_TO`,
containing:

* **Usage details** - date and time in IST and UTC, processing time, session ID,
  client IP, resolved location, and the environment the app ran in.
* **Files uploaded** - file name, format, size, page count, word count and the
  number of highlighted words, per document.
* **Changes highlighted** - total highlighted changes, the split across the two
  documents, matching words, match rate and change rate.
* **Outputs generated** - the names of the two highlighted files.
* **Attachment** - the same record as JSON, for consolidating into a wider MIS.

The mail is sent while the user waits, so the result banner reports honestly whether
it went out. A failure is shown as a warning on screen and never blocks the
comparison. Users can turn reporting off for a session with the sidebar toggle.

Transport: `smtp.gmail.com` over **587/STARTTLS**, falling back to **465/SSL**.
587 is tried first because corporate networks frequently drop the implicit-TLS
connection on 465 — that is exactly what happens on the Bajaj network.

### Credentials

`GMAIL_APP_PASSWORD` is deliberately **not** in any committed file. It is read from
`st.secrets` first, then the environment. `GMAIL_USERNAME` and `MIS_EMAIL_TO` have
non-secret defaults in `mis_mailer.py`, so only the password strictly has to be set.

If the password ever needs rotating, generate a new App Password on the Gmail account
(Google Account -> Security -> 2-Step Verification -> App passwords) and update the
Secrets box. Nothing in the repository needs to change.

## Usage log and the dashboard

Comparisons are also appended to `~/.doc_compare/usage_logs.jsonl`, which backs the
**Usage Dashboard** page.

> On Streamlit Community Cloud that file lives on ephemeral disk and is wiped
> whenever the app restarts or is redeployed. The dashboard will therefore show only
> usage since the last restart. **The MIS e-mails are the durable record.** To make
> the dashboard durable too, point `USAGE_LOG_FILE_PATH` at persistent storage, or
> move the log to a database.

## Debug logging

Engine debug logs echo the text of uploaded documents, so they are off by default.
Set `DOC_COMPARE_DEBUG=1` locally if you need them. Never enable this on a hosted
deployment — the logs would capture the contents of every uploaded file.
