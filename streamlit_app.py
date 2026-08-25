import json
import os
import tempfile
import time
from io import BytesIO
from pathlib import Path

import plotly.graph_objects as go
import streamlit as st


def disable_broken_local_proxy():
    for name in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
        value = os.environ.get(name, "")
        if "127.0.0.1:9" in value or "localhost:9" in value:
            os.environ.pop(name, None)


disable_broken_local_proxy()

from verify_emails import (
    DEFAULT_EMAIL_COLUMN,
    DEFAULT_RESULT_COLUMN,
    DEFAULT_SCORE_COLUMN,
    DEFAULT_SPREADSHEET_URL,
    DEFAULT_WORKSHEET_NAME,
    GOOGLE_CREDENTIALS_ENV,
    ExcelSource,
    GoogleSheetSource,
    clean_email,
    collect_pending,
    column_to_index,
    convert_result,
    index_to_column,
    verify_batch,
)


APP_DIR = Path(".email_verify_app")
SAVED_CREDENTIALS_FILE = APP_DIR / "google_credentials.json"
SAVED_CONFIG_FILE = APP_DIR / "config.json"

st.set_page_config(page_title="EmailVerify Pro", page_icon="@", layout="wide")


def inject_css():
    st.markdown(
        """
<style>
.block-container {
  padding-top: 1.5rem;
  padding-bottom: 3rem;
  max-width: 1180px;
}
.ev-title {
  font-size: 2rem;
  font-weight: 750;
  margin: 0 0 .25rem 0;
}
.ev-subtitle {
  opacity: .75;
  margin-bottom: 1.25rem;
}
.ev-band {
  border: 1px solid rgba(128, 128, 128, .3);
  border-radius: 8px;
  padding: 1rem 1.1rem;
  margin: .75rem 0;
}
.ev-kicker {
  opacity: .65;
  font-size: .85rem;
  text-transform: uppercase;
  letter-spacing: .04em;
  font-weight: 700;
}
.ev-step {
  border-left: 3px solid var(--primary-color, #2563eb);
  padding-left: .9rem;
  margin: .75rem 0;
}
div[data-testid="stMetric"] {
  border: 1px solid rgba(128, 128, 128, .3);
  border-radius: 8px;
  padding: .9rem;
}
div[data-testid="stButton"] > button {
  border-radius: 7px;
  font-weight: 650;
}
</style>
""",
        unsafe_allow_html=True,
    )


def app_header(title, subtitle):
    st.markdown(f'<div class="ev-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="ev-subtitle">{subtitle}</div>', unsafe_allow_html=True)


def load_saved_config():
    try:
        with open(SAVED_CONFIG_FILE, "r", encoding="utf-8") as handle:
            return json.load(handle)
    except Exception:
        return {}


def save_config(config):
    APP_DIR.mkdir(exist_ok=True)
    with open(SAVED_CONFIG_FILE, "w", encoding="utf-8") as handle:
        json.dump(config, handle, indent=2)


def save_credentials(credentials_text):
    credentials = json.loads(credentials_text)
    APP_DIR.mkdir(exist_ok=True)
    with open(SAVED_CREDENTIALS_FILE, "w", encoding="utf-8") as handle:
        json.dump(credentials, handle, indent=2)
    return credentials.get("client_email", "")


def load_credentials_text():
    if SAVED_CREDENTIALS_FILE.exists():
        return SAVED_CREDENTIALS_FILE.read_text(encoding="utf-8")

    try:
        secret_value = st.secrets.get(GOOGLE_CREDENTIALS_ENV, "")
        if secret_value:
            return str(secret_value)
    except Exception:
        pass

    return ""


def remove_saved_credentials():
    if SAVED_CREDENTIALS_FILE.exists():
        SAVED_CREDENTIALS_FILE.unlink()


def get_json_client_email(credentials_text):
    if not credentials_text.strip():
        return ""
    try:
        return json.loads(credentials_text).get("client_email", "")
    except Exception:
        return ""


def unique_email_map(pending):
    email_to_rows = {}
    for item in pending:
        email_to_rows.setdefault(item["email"], []).append(item["row"])
    return email_to_rows


def build_excel_source(uploaded_file, worksheet_name):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    return ExcelSource(tmp_path, worksheet_name), tmp_path


@st.cache_resource(show_spinner="Connecting to Google Sheet...")
def cached_google_worksheet(spreadsheet_url, worksheet_name, credentials_text):
    if credentials_text.strip():
        json.loads(credentials_text)
        os.environ[GOOGLE_CREDENTIALS_ENV] = credentials_text.strip()
    return GoogleSheetSource(spreadsheet_url, worksheet_name, "credentials.json")


@st.cache_data(ttl=60, show_spinner="Reading sheet rows...")
def cached_google_rows(spreadsheet_url, worksheet_name, credentials_text, cache_bust):
    source = cached_google_worksheet(spreadsheet_url, worksheet_name, credentials_text)
    return source.read_rows()


def make_source(config):
    if config["mode"] == "Google Sheets":
        credentials_text = config.get("credentials_text", "") or load_credentials_text()
        return cached_google_worksheet(
            config["spreadsheet_url"],
            config["worksheet_name"],
            credentials_text.strip(),
        ), None

    uploaded_file = st.session_state.get("excel_file")
    if uploaded_file is None:
        raise ValueError("Upload an Excel file in Setup first.")
    return build_excel_source(uploaded_file, config.get("excel_sheet", ""))


def read_rows_for(source, config):
    if config["mode"] == "Google Sheets":
        credentials_text = (config.get("credentials_text", "") or load_credentials_text()).strip()
        cache_bust = st.session_state.get("sheet_cache_bust", 0)
        return cached_google_rows(
            config["spreadsheet_url"], config["worksheet_name"], credentials_text, cache_bust
        )
    return source.read_rows()


def scan_source(rows, config):
    email_col = column_to_index(config["email_column"])
    result_col = column_to_index(config["result_column"])

    start_row = int(config["start_row"])
    end_row = min(int(config["end_row"]), len(rows))
    if end_row < start_row:
        raise ValueError("End row must be greater than or equal to start row.")

    pending, skipped, blank, _ = collect_pending(
        rows,
        start_row,
        end_row,
        email_col,
        result_col,
        set(),
    )
    email_to_rows = unique_email_map(pending)

    return {
        "rows": rows,
        "start_row": start_row,
        "end_row": end_row,
        "selected_rows": max(0, end_row - start_row + 1),
        "pending": pending,
        "email_to_rows": email_to_rows,
        "unique_emails": list(email_to_rows),
        "already_done": skipped,
        "blank": blank,
        "email_col": email_col,
        "result_col": result_col,
        "score_col": column_to_index(config["score_column"]),
    }


def _stacked_bar(segments, height=110):
    fig = go.Figure()
    total = max(sum(value for _, value, _ in segments), 1)
    for label, value, color in segments:
        fig.add_bar(
            y=[""],
            x=[value],
            name=f"{label} ({value})",
            orientation="h",
            marker_color=color,
            text=str(value) if value else "",
            textposition="inside",
            insidetextanchor="middle",
        )
    fig.update_layout(
        barmode="stack",
        height=height,
        margin=dict(l=0, r=0, t=8, b=8),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="left", x=0),
        xaxis=dict(visible=False, range=[0, total]),
        yaxis=dict(visible=False),
    )
    return fig


def render_composition_chart(container, scan):
    with container.container():
        st.caption("Selected rows")
        segments = [
            ("Verified", scan["already_done"], "#2a78d6"),
            ("Pending", len(scan["pending"]), "#eb6834"),
            ("Blank", scan["blank"], "#898781"),
        ]
        if sum(v for _, v, _ in segments) == 0:
            st.caption("No rows in the selected range.")
            return
        st.plotly_chart(_stacked_bar(segments), use_container_width=True, theme="streamlit")


def render_outcome_chart(container, valid, invalid):
    with container.container():
        st.caption("This run: valid vs invalid")
        if valid + invalid == 0:
            st.caption("Run verification to see the breakdown here.")
            return
        segments = [
            ("Valid", valid, "#0ca30c"),
            ("Invalid", invalid, "#d03b3b"),
        ]
        st.plotly_chart(_stacked_bar(segments), use_container_width=True, theme="streamlit")


def render_kpis(container, scan, done, valid, invalid, total_unique):
    with container.container():
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Rows selected", scan["selected_rows"])
        m2.metric("Emails waiting", len(scan["pending"]))
        m3.metric("Unique emails", total_unique)
        m4.metric("Already done", scan["already_done"])
        m5.metric("Blank emails", scan["blank"])

        r1, r2, r3 = st.columns(3)
        r1.metric("Happened now", done)
        r2.metric("Valid now", valid)
        r3.metric("Invalid now", invalid)

        st.progress(0 if not total_unique else min(done / total_unique, 1.0))


def sidebar_nav():
    st.sidebar.markdown("## EmailVerify Pro")
    st.sidebar.caption("Slow, one-by-one email verification")

    if "page" not in st.session_state:
        st.session_state.page = "Setup"

    page = st.sidebar.radio(
        "Navigation",
        ["Setup", "Dashboard", "Instructions"],
        index=["Setup", "Dashboard", "Instructions"].index(st.session_state.page),
    )
    st.session_state.page = page

    st.sidebar.divider()
    st.sidebar.caption("Current output")
    st.sidebar.write(f"Email: `{DEFAULT_EMAIL_COLUMN}`")
    st.sidebar.write(f"Result: `{DEFAULT_RESULT_COLUMN}`")
    st.sidebar.write(f"Score: `{DEFAULT_SCORE_COLUMN}`")

    return page


def setup_page():
    app_header("Setup", "Connect a Google Sheet or upload Excel, then open the dashboard.")
    saved = load_saved_config()

    with st.container(border=True):
        st.markdown('<div class="ev-kicker">Data Source</div>', unsafe_allow_html=True)
        mode = st.segmented_control(
            "Source",
            ["Google Sheets", "Excel"],
            default=saved.get("mode", "Google Sheets"),
            label_visibility="collapsed",
        )

        batch_col, coming_col = st.columns([1, 2])
        batch_col.toggle("Verify 1 by 1", value=True, disabled=True)
        coming_col.button("Batch mode coming soon", disabled=True, use_container_width=True)

    with st.container(border=True):
        st.markdown('<div class="ev-kicker">Columns and Range</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        email_column = c1.text_input("Email column", saved.get("email_column", DEFAULT_EMAIL_COLUMN))
        result_column = c2.text_input("Result column", saved.get("result_column", DEFAULT_RESULT_COLUMN))
        score_column = c3.text_input("Score / reason column", saved.get("score_column", DEFAULT_SCORE_COLUMN))

        r1, r2, r3 = st.columns(3)
        start_row = r1.number_input("Start row", min_value=1, value=int(saved.get("start_row", 2)), step=1)
        end_row = r2.number_input("End row", min_value=1, value=int(saved.get("end_row", 100)), step=1)
        delay_seconds = r3.number_input(
            "Delay after each email",
            min_value=0,
            value=int(saved.get("delay_seconds", 8)),
            step=1,
        )

    spreadsheet_url = saved.get("spreadsheet_url", DEFAULT_SPREADSHEET_URL)
    worksheet_name = saved.get("worksheet_name", DEFAULT_WORKSHEET_NAME)
    credentials_text = ""
    excel_sheet = saved.get("excel_sheet", "")

    if mode == "Google Sheets":
        with st.container(border=True):
            st.markdown('<div class="ev-kicker">Google Sheet Access</div>', unsafe_allow_html=True)
            spreadsheet_url = st.text_input("Spreadsheet URL", spreadsheet_url)
            worksheet_name = st.text_input("Worksheet / tab name", worksheet_name)

            existing_credentials = load_credentials_text()
            client_email = get_json_client_email(existing_credentials)
            if client_email:
                st.success(f"Saved Google JSON found. Share your sheet with: {client_email}")

            credentials_text = st.text_area(
                "Google service account JSON",
                value="",
                height=150,
                placeholder="Paste full JSON here. You can save it on this computer.",
            )

            s1, s2 = st.columns(2)
            remember_json = s1.checkbox("Save this JSON on this computer", value=False)
            if s2.button("Forget saved JSON", use_container_width=True):
                remove_saved_credentials()
                st.success("Saved JSON removed.")
                st.rerun()
    else:
        with st.container(border=True):
            st.markdown('<div class="ev-kicker">Excel Upload</div>', unsafe_allow_html=True)
            uploaded_file = st.file_uploader("Upload .xlsx file", type=["xlsx"])
            if uploaded_file:
                st.session_state.excel_file = uploaded_file
            excel_sheet = st.text_input("Worksheet name", excel_sheet)
            st.caption("Leave worksheet blank to use the first sheet.")
            remember_json = False

    config = {
        "mode": mode,
        "spreadsheet_url": spreadsheet_url,
        "worksheet_name": worksheet_name,
        "excel_sheet": excel_sheet.strip(),
        "email_column": email_column.strip() or DEFAULT_EMAIL_COLUMN,
        "result_column": result_column.strip() or DEFAULT_RESULT_COLUMN,
        "score_column": score_column.strip() or DEFAULT_SCORE_COLUMN,
        "start_row": int(start_row),
        "end_row": int(end_row),
        "delay_seconds": int(delay_seconds),
        "retry_delay": 30,
        "jitter": 3,
    }

    st.info(
        f"Emails read from column {config['email_column'].upper()}; "
        f"results write to {config['result_column'].upper()} and {config['score_column'].upper()}."
    )

    c1, c2 = st.columns([1, 1])
    if c1.button("Save setup", type="primary", use_container_width=True):
        if mode == "Google Sheets" and credentials_text.strip() and remember_json:
            client_email = save_credentials(credentials_text)
            st.success(f"Google JSON saved. Share your sheet with: {client_email}")

        if mode == "Google Sheets" and credentials_text.strip():
            json.loads(credentials_text)
            config["credentials_text"] = credentials_text.strip()

        save_config({k: v for k, v in config.items() if k != "credentials_text"})
        st.session_state.config = config
        st.success("Setup saved.")

    if c2.button("Save and open dashboard", use_container_width=True):
        if mode == "Google Sheets" and credentials_text.strip() and remember_json:
            save_credentials(credentials_text)
        if mode == "Google Sheets" and credentials_text.strip():
            json.loads(credentials_text)
            config["credentials_text"] = credentials_text.strip()

        save_config({k: v for k, v in config.items() if k != "credentials_text"})
        st.session_state.config = config
        st.session_state.page = "Dashboard"
        st.rerun()


def dashboard_page():
    top_col, refresh_col = st.columns([5, 1])
    with top_col:
        app_header("Dashboard", "Review volume, then verify emails one at a time.")
    with refresh_col:
        if st.button("Refresh sheet data", use_container_width=True):
            st.session_state.sheet_cache_bust = st.session_state.get("sheet_cache_bust", 0) + 1
            st.rerun()

    config = st.session_state.get("config") or load_saved_config()
    if not config:
        st.warning("Complete Setup first.")
        if st.button("Go to Setup"):
            st.session_state.page = "Setup"
            st.rerun()
        return

    try:
        source, excel_tmp_path = make_source(config)
        rows = read_rows_for(source, config)
        scan = scan_source(rows, config)
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    total_unique = len(scan["unique_emails"])
    done_this_run = st.session_state.get("done_this_run", 0)
    valid_this_run = st.session_state.get("valid_this_run", 0)
    invalid_this_run = st.session_state.get("invalid_this_run", 0)

    kpi_slot = st.empty()
    render_kpis(kpi_slot, scan, done_this_run, valid_this_run, invalid_this_run, total_unique)

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        composition_slot = st.empty()
        render_composition_chart(composition_slot, scan)
    with chart_col2:
        outcome_slot = st.empty()
        render_outcome_chart(outcome_slot, valid_this_run, invalid_this_run)

    with st.container(border=True):
        st.markdown('<div class="ev-kicker">Run Control</div>', unsafe_allow_html=True)
        st.write(
            f"Mode: **1 by 1** | Delay: **{config.get('delay_seconds', 8)} seconds** after each email | "
            f"Columns: **{index_to_column(scan['email_col'])} -> {index_to_column(scan['result_col'])}, "
            f"{index_to_column(scan['score_col'])}**"
        )
        run_limit = st.number_input(
            "How many emails to verify now",
            min_value=1,
            max_value=max(total_unique, 1),
            value=min(total_unique, 25) if total_unique else 1,
            step=1,
        )
        start = st.button("Start verifying 1 by 1", type="primary", disabled=total_unique == 0)

    status_box = st.empty()
    log_box = st.empty()
    run_progress = st.empty()

    if start:
        logs = []
        verified = 0
        for email in scan["unique_emails"][: int(run_limit)]:
            row_numbers = scan["email_to_rows"][email]
            status_box.info(f"Verifying {email}")

            try:
                results = verify_batch([email], config.get("retry_delay", 30), config.get("jitter", 3))
                result = results[0] if results else {"email": email, "status": "FAILED"}
                verification, score_reason = convert_result(result)

                row_updates = [
                    (row_number, verification, score_reason)
                    for row_number in row_numbers
                ]
                source.write_results(row_updates, scan["result_col"], scan["score_col"])
                source.save()

                done_this_run += len(row_updates)
                if verification == "Valid":
                    valid_this_run += len(row_updates)
                else:
                    invalid_this_run += len(row_updates)

                st.session_state.done_this_run = done_this_run
                st.session_state.valid_this_run = valid_this_run
                st.session_state.invalid_this_run = invalid_this_run

                rows_text = ", ".join(str(row) for row in row_numbers)
                logs.append(f"Rows {rows_text}: {email} -> {verification} | {score_reason}")
            except Exception as exc:
                logs.append(f"{email} -> ERROR | {exc}")

            verified += 1
            log_box.code("\n".join(logs[-80:]), language="text")
            run_progress.progress(min(verified / int(run_limit), 1.0))
            render_kpis(kpi_slot, scan, done_this_run, valid_this_run, invalid_this_run, total_unique)
            render_outcome_chart(outcome_slot, valid_this_run, invalid_this_run)

            if verified < int(run_limit):
                delay = int(config.get("delay_seconds", 8))
                status_box.info(f"Saved. Waiting {delay}s before next email...")
                time.sleep(delay)

        status_box.success(f"Finished {verified} unique emails.")
        st.session_state.sheet_cache_bust = st.session_state.get("sheet_cache_bust", 0) + 1

        if config["mode"] == "Excel":
            output = BytesIO()
            source.workbook.save(output)
            st.download_button(
                "Download verified Excel",
                data=output.getvalue(),
                file_name="verified_emails.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    if config["mode"] == "Excel" and excel_tmp_path:
        try:
            os.remove(excel_tmp_path)
        except OSError:
            pass


def instructions_page():
    app_header("Instructions", "A new-user setup guide with direct credential links.")

    with st.container(border=True):
        st.markdown('<div class="ev-kicker">Columns</div>', unsafe_allow_html=True)
        st.markdown(
            """
| Column | Use |
| --- | --- |
| C | Email address |
| W | Valid / Invalid result |
| X | Score and reason |
"""
        )

    with st.container(border=True):
        st.markdown('<div class="ev-kicker">Direct Links</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        c1.link_button("Create Google Project", "https://console.cloud.google.com/projectcreate", use_container_width=True)
        c2.link_button("Enable Sheets API", "https://console.cloud.google.com/apis/library/sheets.googleapis.com", use_container_width=True)
        c3.link_button("Enable Drive API", "https://console.cloud.google.com/apis/library/drive.googleapis.com", use_container_width=True)
        c4, c5, c6 = st.columns(3)
        c4.link_button("Create Service Account", "https://console.cloud.google.com/iam-admin/serviceaccounts/create", use_container_width=True)
        c5.link_button("Service Accounts", "https://console.cloud.google.com/iam-admin/serviceaccounts", use_container_width=True)
        c6.link_button("Open Google Sheets", "https://sheets.google.com", use_container_width=True)

    with st.container(border=True):
        st.markdown('<div class="ev-kicker">Get Google JSON</div>', unsafe_allow_html=True)
        st.markdown(
            """
<div class="ev-step">1. Create/select a Google Cloud project.</div>
<div class="ev-step">2. Enable Google Sheets API and Google Drive API.</div>
<div class="ev-step">3. Create a service account named <code>email-verify</code>.</div>
<div class="ev-step">4. Open that service account, go to <b>Keys</b>, click <b>Add key</b>, choose <b>Create new key</b>, then select <b>JSON</b>.</div>
<div class="ev-step">5. Open the downloaded JSON and copy the whole file into the Setup page.</div>
""",
            unsafe_allow_html=True,
        )
        st.warning("Google lets you download the private JSON key only once. Keep it private.")

    with st.container(border=True):
        st.markdown('<div class="ev-kicker">Share Your Sheet With The JSON Email</div>', unsafe_allow_html=True)
        st.markdown(
            """
1. Open your downloaded JSON file.
2. Find this field:
"""
        )
        st.code('"client_email": "email-verify@your-project.iam.gserviceaccount.com"', language="json")
        st.markdown(
            """
3. Copy only the email address inside `client_email`.
4. Open your Google Sheet.
5. Click **Share**.
6. Paste that `client_email`.
7. Set permission to **Editor**.
8. Click **Send**.
"""
        )

    with st.container(border=True):
        st.markdown('<div class="ev-kicker">First Safe Run</div>', unsafe_allow_html=True)
        st.code(
            """Source: Google Sheets
Email column: C
Result column: W
Score / reason column: X
Start row: 2
End row: 10
Delay after each email: 8 to 15 seconds""",
            language="text",
        )


def main():
    inject_css()
    page = sidebar_nav()

    if page == "Setup":
        setup_page()
    elif page == "Dashboard":
        dashboard_page()
    else:
        instructions_page()


if __name__ == "__main__":
    main()
