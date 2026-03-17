"""
Seamfix Financial Intelligence Suite
Streamlit Cloud deployment — serves 4 interactive dashboards.

Data sources:
  • Revenue data:  Google Sheet (live) or uploaded xlsx
  • Cash reports:  uploaded xlsx files or pre-loaded data/
  • Budget:        uploaded xlsx or pre-loaded data/
"""

import streamlit as st
import streamlit.components.v1 as components
import os, sys, shutil, tempfile, glob, io, urllib.request
from pathlib import Path

# ── Page config ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="Seamfix Financial Intelligence",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Constants ────────────────────────────────────────────────────────
APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
GENERATED_DIR = APP_DIR / "generated"
GENERATED_DIR.mkdir(exist_ok=True)

GOOGLE_SHEET_ID = "1XKIE9eRP8H1AWpuMAJA0U8bM7pQ9o1jvoQobc6aUn5s"
GOOGLE_DRIVE_FOLDER_ID = "1vLq8m030d1ifL6nAVuo9LT5N9NSeGs9U"
REVENUE_FILENAME = "2026 Path to Revenue (1).xlsx"
BUDGET_FILENAME = "2026 LEAN BUDGET.xlsx"

DASHBOARDS = {
    "Cash Overview": {
        "icon": "💰",
        "script": "generate_dashboard.py",
        "output": "dashboard.html",
        "description": "Weekly cash position, inflows/outflows, FX rates, and bank balances"
    },
    "Expense & Vendor Analysis": {
        "icon": "📊",
        "script": "generate_expense_dashboard.py",
        "output": "expense_dashboard.html",
        "description": "Category breakdown, vendor analysis, investment tracking"
    },
    "Budget vs Actual": {
        "icon": "📋",
        "script": "generate_budget_dashboard.py",
        "output": "budget_dashboard.html",
        "description": "₦5.1B annual budget mapped against actual weekly spend"
    },
    "Revenue & Fundability": {
        "icon": "🚀",
        "script": "generate_revenue_dashboard.py",
        "output": "revenue_dashboard.html",
        "description": "$10M revenue pipeline, fundability analysis, critical actions"
    },
}


# ── Helper: Google OAuth check ───────────────────────────────────────
def check_auth():
    """
    Check if user is authenticated via Google OAuth using st.login().
    Falls back to open access if OAuth is not configured.
    """
    # If running locally or auth disabled, allow access
    if os.getenv("DISABLE_AUTH", "false").lower() == "true":
        return True

    # Check if OAuth is fully configured in secrets
    auth_conf = st.secrets.get("auth", {})
    has_oauth = all(
        k in auth_conf
        for k in ["client_id", "client_secret", "redirect_uri", "cookie_secret", "server_metadata_url"]
    )

    if has_oauth:
        try:
            # OAuth is configured — require login
            user = st.user
            if not user.is_logged_in:
                # Show login page
                st.markdown(
                    """
                    <div style="text-align:center;padding:80px 20px">
                        <h1 style="color:#00D4AA;margin-bottom:8px">Seamfix Financial Intelligence</h1>
                        <p style="color:#94a3b8;margin-bottom:40px">Sign in with your authorized Google account to access the dashboards.</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.login("google")
                st.stop()
                return False

            # User is logged in — check against allowlist
            email = user.email.lower().strip()
            allowed_emails = [
                e.lower().strip()
                for e in auth_conf.get("allowed_emails", [])
            ]

            if allowed_emails and email not in allowed_emails:
                st.error(
                    f"Access denied. **{email}** is not on the authorized users list.\n\n"
                    "Contact the dashboard administrator to request access."
                )
                if st.button("Sign out"):
                    st.logout()
                st.stop()
                return False

            return True
        except Exception as e:
            # OAuth misconfigured — show warning and allow access for now
            st.sidebar.warning(f"OAuth not fully configured: {e}")
            return True

    # No OAuth configured — allow access (development mode)
    return True


# ── Helper: fetch Google Sheet as xlsx ───────────────────────────────
@st.cache_data(ttl=300, show_spinner="Fetching live revenue data from Google Sheet...")
def fetch_google_sheet_xlsx(sheet_id):
    """Download Google Sheet as xlsx. Requires sheet to be shared
    with 'Anyone with the link' (Viewer) or a service account."""
    errors = []

    try:
        # Try public export
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        resp = urllib.request.urlopen(req, timeout=30)
        return resp.read()
    except Exception as e:
        errors.append(f"Public access: {e}")

    # Try with service account (if configured in secrets)
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets.readonly",
                "https://www.googleapis.com/auth/drive.readonly",
            ],
        )
        gc = gspread.authorize(creds)
        spreadsheet = gc.open_by_key(sheet_id)
        return spreadsheet.export(format=gspread.utils.ExportFormat.EXCEL)
    except KeyError:
        errors.append("Service account: No gcp_service_account in secrets")
    except Exception as e:
        errors.append(f"Service account: {e}")

    # Store errors for debugging display
    st.session_state["_gsheet_errors"] = errors
    return None


# ── Helper: fetch cash reports from Google Drive folder ──────────────
@st.cache_data(ttl=300, show_spinner="Fetching cash reports from Google Drive...")
def fetch_drive_folder_files(folder_id):
    """
    List and download xlsx files from a Google Drive folder.
    Requires a service account with access to the folder.
    Returns dict of {filename: bytes} or None.
    """
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload

        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/drive.readonly"],
        )
        service = build("drive", "v3", credentials=creds)

        # List xlsx files in folder
        query = f"'{folder_id}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
        results = service.files().list(
            q=query, fields="files(id, name)", pageSize=50
        ).execute()
        files = results.get("files", [])

        if not files:
            st.session_state["_gdrive_errors"] = ["No xlsx files found in folder (folder may not be shared with service account)"]
            return None

        # Download each file
        downloaded = {}
        for f in files:
            request = service.files().get_media(fileId=f["id"])
            buf = io.BytesIO()
            downloader = MediaIoBaseDownload(buf, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            downloaded[f["name"]] = buf.getvalue()

        return downloaded
    except KeyError:
        st.session_state["_gdrive_errors"] = ["No gcp_service_account in secrets"]
        return None
    except Exception as e:
        st.session_state["_gdrive_errors"] = [str(e)]
        return None


# ── Helper: prepare data folder ─────────────────────────────────────
def prepare_data_folder():
    """
    Build a temp folder with all the data files needed by generators.
    Merges: pre-loaded data/ + Google Drive files + Google Sheet + uploads.
    Returns path to the folder.
    """
    data_path = GENERATED_DIR / "data_working"
    data_path.mkdir(exist_ok=True)

    # 1. Copy pre-loaded data files (bundled baseline)
    if DATA_DIR.exists():
        for f in DATA_DIR.glob("*.xlsx"):
            dest = data_path / f.name
            if not dest.exists():
                shutil.copy2(f, dest)

    # 2. Fetch cash reports from Google Drive folder (overwrites bundled)
    if GOOGLE_DRIVE_FOLDER_ID:
        drive_files = fetch_drive_folder_files(GOOGLE_DRIVE_FOLDER_ID)
        if drive_files:
            for fname, fbytes in drive_files.items():
                (data_path / fname).write_bytes(fbytes)
            st.session_state["cash_source"] = f"Google Drive ({len(drive_files)} files)"
        else:
            bundled = len(list(data_path.glob("Cash Report*.xlsx")))
            st.session_state["cash_source"] = f"Local files ({bundled} bundled)" if bundled else "Not available"

    # 3. Try fetching live revenue data from Google Sheet
    if GOOGLE_SHEET_ID:
        xlsx_bytes = fetch_google_sheet_xlsx(GOOGLE_SHEET_ID)
        if xlsx_bytes:
            (data_path / REVENUE_FILENAME).write_bytes(xlsx_bytes)
            st.session_state["revenue_source"] = "Google Sheet (live)"
        else:
            if (data_path / REVENUE_FILENAME).exists():
                st.session_state["revenue_source"] = "Local file (bundled)"
            else:
                st.session_state["revenue_source"] = "Not available"

    # 4. Handle uploaded files (from sidebar — overwrites everything)
    if "uploaded_files" in st.session_state:
        for uploaded in st.session_state.uploaded_files:
            dest = data_path / uploaded.name
            dest.write_bytes(uploaded.getvalue())

    return str(data_path)


# ── Helper: generate a dashboard ────────────────────────────────────
def generate_dashboard(script_name, data_folder, output_name):
    """Run a generator script and return the HTML content."""
    output_path = GENERATED_DIR / output_name

    # Import and run the generator
    script_path = APP_DIR / script_name
    if not script_path.exists():
        return None

    # Run as subprocess to avoid import conflicts
    import subprocess

    result = subprocess.run(
        [sys.executable, str(script_path), data_folder],
        capture_output=True,
        text=True,
        timeout=60,
    )

    if result.returncode != 0:
        st.error(f"Generator error: {result.stderr[-500:]}")
        return None

    # The generators write to data_folder/output_name
    generated_file = Path(data_folder) / output_name
    if generated_file.exists():
        return generated_file.read_text(encoding="utf-8")

    # Some generators also write to a secondary location
    alt_path = Path(data_folder).parent / "outputs" / output_name
    if alt_path.exists():
        return alt_path.read_text(encoding="utf-8")

    return None


# ── Helper: fix nav links for embedded view ──────────────────────────
def fix_html_for_streamlit(html_content):
    """
    Remove the inter-dashboard nav bar from embedded HTML since
    Streamlit sidebar handles navigation. Also ensure proper sizing.
    """
    if not html_content:
        return html_content

    # Remove the nav bar div (it links to file:// paths that won't work)
    import re

    html_content = re.sub(
        r'<div class="nav-bar">.*?</div>',
        "",
        html_content,
        flags=re.DOTALL,
    )

    # Add responsive height to body
    html_content = html_content.replace(
        "padding:24px;min-height:100vh",
        "padding:24px;min-height:100vh;width:100%",
    )

    return html_content


# ── Main App ─────────────────────────────────────────────────────────
def main():
    # Auth check
    check_auth()

    # ── Sidebar ──────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(
            """
            <div style="text-align:center;padding:10px 0 20px">
                <h2 style="margin:0;color:#00D4AA">Seamfix</h2>
                <p style="margin:0;color:#94a3b8;font-size:0.85em">Financial Intelligence Suite</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Dashboard selector
        selected = st.radio(
            "Dashboard",
            list(DASHBOARDS.keys()),
            format_func=lambda x: f"{DASHBOARDS[x]['icon']}  {x}",
            label_visibility="collapsed",
        )

        st.divider()

        # Data management section
        st.markdown("##### Data Management")

        # File uploader for new cash reports
        uploaded = st.file_uploader(
            "Upload new cash reports",
            type=["xlsx"],
            accept_multiple_files=True,
            key="uploaded_files",
            help="Drop new weekly cash report xlsx files here",
        )

        # Regenerate button
        if st.button("🔄 Regenerate Dashboards", use_container_width=True):
            # Clear cached data
            st.cache_data.clear()
            for key in list(st.session_state.keys()):
                if key.startswith("html_"):
                    del st.session_state[key]
            st.rerun()

        st.divider()

        # Data source status
        st.markdown("##### Data Sources")
        rev_source = st.session_state.get("revenue_source", "Checking...")
        cash_source = st.session_state.get("cash_source", "Checking...")
        upload_count = len(uploaded) if uploaded else 0

        st.caption(f"📈 Revenue: {rev_source}")
        st.caption(f"💵 Cash reports: {cash_source}" + (f" + {upload_count} uploaded" if upload_count else ""))
        st.caption(f"📋 Budget: {'Available' if (DATA_DIR / BUDGET_FILENAME).exists() else 'Missing'}")

        # Debug info for connection issues
        gsheet_errs = st.session_state.get("_gsheet_errors", [])
        gdrive_errs = st.session_state.get("_gdrive_errors", [])
        if gsheet_errs or gdrive_errs:
            with st.expander("Connection issues"):
                if gsheet_errs:
                    st.caption("Google Sheet:")
                    for e in gsheet_errs:
                        st.caption(f"  {e}")
                if gdrive_errs:
                    st.caption("Google Drive:")
                    for e in gdrive_errs:
                        st.caption(f"  {e}")

        # Show logged-in user and logout button
        try:
            user = st.user
            if user.is_logged_in:
                st.divider()
                st.caption(f"Signed in as **{user.email}**")
                st.logout("Sign out")
        except Exception:
            pass

    # ── Main content ─────────────────────────────────────────────────
    dash = DASHBOARDS[selected]

    # Prepare data
    data_folder = prepare_data_folder()

    # Generate or use cached HTML
    cache_key = f"html_{selected}"
    if cache_key not in st.session_state:
        with st.spinner(f"Generating {selected} dashboard..."):
            html = generate_dashboard(
                dash["script"], data_folder, dash["output"]
            )
            if html:
                html = fix_html_for_streamlit(html)
                st.session_state[cache_key] = html
            else:
                st.session_state[cache_key] = None

    html_content = st.session_state.get(cache_key)

    if html_content:
        # Render the dashboard
        components.html(html_content, height=4000, scrolling=True)
    else:
        st.warning(
            f"Could not generate the {selected} dashboard. "
            "Check that all required data files are available."
        )
        st.info(
            f"This dashboard needs: "
            + ("Cash Report xlsx files" if "Cash" in selected or "Expense" in selected else "")
            + ("Revenue xlsx + Budget xlsx" if "Revenue" in selected or "Budget" in selected else "")
        )


if __name__ == "__main__":
    main()
