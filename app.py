#!/usr/bin/env python
# coding: utf-8

import streamlit as st
import pandas as pd
import yagmail
import re
import tempfile
import os
import base64
from io import StringIO
from streamlit_quill import st_quill
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter

# -------------------------------------------------------
# EXCEL LOGGING FUNCTION (CLEAN FORMAT)
# -------------------------------------------------------
def export_logs_excel(logs):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"email_logs_{ts}.xlsx"
    filepath = f"/tmp/{filename}"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Email Logs"

    headers = ["Row", "Email", "Status", "Details", "Timestamp"]
    ws.append(headers)

    # Bold headers
    for col in range(1, len(headers) + 1):
        ws.cell(row=1, column=col).font = Font(bold=True)

    # Insert log rows
    for row in logs:
        ws.append(row)

    # Borders and auto-width
    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin

    # Auto column width
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 2

    wb.save(filepath)
    return filepath, filename


# -------------------------------------------------------
# CONFIG & STYLING
# -------------------------------------------------------
st.set_page_config(page_title="Email Blaster", layout="centered")

def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def set_png_as_page_bg(png_file):
    bin_str = get_base64_of_bin_file(png_file)
    page_bg_img = f'''
    <style>
    .stApp {{
        background-image: url("data:image/jpg;base64,{bin_str}");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
    }}
    
    /* Card Container - Floating Effect */
    .block-container {{
        background-color: rgba(255, 255, 255, 0.95);
        padding: 2rem; /* Reduced padding */
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        margin: 3rem auto; /* Reduced vertical spacing */
        max-width: 800px;
    }}

    /* Force Dark Text for Visibility */
    h1, h2, h3, h4, h5, h6, p, label, li, span, div {{
        color: #2c3e50 !important;
    }}
    
    /* Input Fields */
    .stTextInput > div > div > input {{
        background-color: #f8f9fa;
        border: 1px solid #ced4da;
        border-radius: 8px;
        color: #495057 !important;
    }}
    .stTextInput > div > div > input:focus {{
        border-color: #4dabf7;
        box-shadow: 0 0 0 3px rgba(77, 171, 247, 0.2);
    }}

    /* File Uploader Styling */
    [data-testid='stFileUploader'] {{
        background-color: #f1f3f5;
        border: 2px dashed #ced4da;
        border-radius: 10px;
        padding: 1rem;
        transition: border-color 0.3s;
    }}
    [data-testid='stFileUploader']:hover {{
        border-color: #4dabf7;
    }}
    [data-testid='stFileUploader'] section {{
        background-color: transparent !important;
    }}
    [data-testid='stFileUploader'] button {{
        background-color: #ffffff !important;
        color: #3b82f6 !important;
        border: 1px solid #3b82f6 !important;
    }}
    [data-testid='stFileUploader'] button:hover {{
        background-color: #eff6ff !important;
    }}
    /* Fix text inside uploader */
    [data-testid='stFileUploader'] small, [data-testid='stFileUploader'] span {{
        color: #6c757d !important;
    }}

    /* Buttons */
    .stButton > button {{
        background-color: #3b82f6;
        color: white !important;
        border-radius: 8px;
        border: none;
        padding: 0.6rem 1.2rem;
        font-weight: 600;
        width: 100%;
        transition: all 0.3s ease;
    }}
    .stButton > button:hover {{
        background-color: #2563eb;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
        transform: translateY(-1px);
    }}
    
    /* Headers specific styling */
    h1 {{
        font-weight: 800;
        color: #1a202c !important;
        margin-bottom: 1.5rem;
    }}
    
    h2 {{
        font-weight: 600;
        margin-top: 2rem;
        border-bottom: 2px solid #f0f0f0;
        padding-bottom: 0.5rem;
    }}

    /* Fix for Streamlit's default top padding */
    .main .block-container {{
        padding-top: 2rem;
        padding-bottom: 2rem;
    }}

    /* Hide Streamlit Branding (Footer, Header, Toolbar, etc.) */
    #MainMenu {{display: none !important;}}
    footer {{display: none !important;}}
    header {{display: none !important;}}
    [data-testid="stToolbar"] {{display: none !important;}}
    [data-testid="stDecoration"] {{display: none !important;}}
    [data-testid="stFooter"] {{display: none !important;}}
    [data-testid="stStatusWidget"] {{display: none !important;}}
    [data-testid="stHeader"] {{display: none !important;}}
    .stDeployButton {{display: none !important;}}
    </style>
    '''
    st.markdown(page_bg_img, unsafe_allow_html=True)

try:
    set_png_as_page_bg('UIDBC.jpg')
except Exception as e:
    st.warning(f"Could not load background image: {e}")

st.title("📧 Email Blaster")

st.markdown("""
Upload an Excel (.xlsx) with an **email** column.  
Use placeholders like `{name}`, `{position}`, `{company}`, etc.
""")


# -------------------------------------------------------
# REMEMBER ME SYSTEM (Option B)
# -------------------------------------------------------

# Initialize states
if "saved_email" not in st.session_state:
    st.session_state.saved_email = None

if "saved_pass" not in st.session_state:
    st.session_state.saved_pass = None

if "remember_email" not in st.session_state:
    st.session_state.remember_email = False

if "remember_pass_session" not in st.session_state:
    st.session_state.remember_pass_session = False

# Load email from browser URL params
# UPDATED: Use st.query_params instead of experimental_get_query_params
if "email" in st.query_params and st.session_state.saved_email is None:
    st.session_state.saved_email = st.query_params["email"]


# -------------------------------------------------------
# FIELD DETECTION
# -------------------------------------------------------
FIELD_MAP = {
    "email": ["email", "mail", "e-mail", "emailaddress", "emailid"],
    "name": ["name", "fullname", "full name", "nama", "nama lengkap"],
    "company": ["company", "organization", "org", "perusahaan", "instansi"],
    "position": ["position", "jobtitle", "title", "jabatan", "role"]
}

def normalize(text):
    return re.sub(r"[^a-z0-9]", "", text.lower())

def detect_columns(df):
    detected = {}
    normalized_cols = {normalize(c): c for c in df.columns}

    for field, aliases in FIELD_MAP.items():

        # Hard match for email
        if field == "email":
            for norm, real in normalized_cols.items():
                if norm == "email":
                    detected[field] = real
                    break
            if field in detected:
                continue

        # Alias matching
        for alias in aliases:
            alias_norm = normalize(alias)
            for norm, real in normalized_cols.items():
                if alias_norm in norm or norm in alias_norm:
                    detected[field] = real
                    break
            if field in detected:
                break

    return detected

# -------------------------------------------------------
# QUILL SETUP
# -------------------------------------------------------
if "body_html" not in st.session_state:
    st.session_state.body_html = ""

if "quill_initialized" not in st.session_state:
    st.session_state.quill_initialized = False


# -------------------------------------------------------
# EMAIL LOGIN
# -------------------------------------------------------
st.header("1. Email Account Login")

# Remember Me checkboxes
st.session_state.remember_email = st.checkbox("Remember Email")
st.session_state.remember_pass_session = st.checkbox("Remember App Password (this session only)")

# Email field (load if saved)
email_user = st.text_input(
    "Your Email Address",
    placeholder="example@gmail.com",
    value=st.session_state.saved_email if st.session_state.saved_email else ""
)

# App password field (session only)
email_pass = st.text_input(
    "App Password (NOT your regular password)",
    type="password",
    value=st.session_state.saved_pass if st.session_state.remember_pass_session else ""
)

st.info("For Gmail: Create an App Password at https://myaccount.google.com/apppasswords")

# Save email if remembered
if st.session_state.remember_email and email_user:
    st.session_state.saved_email = email_user
    st.query_params["email"] = email_user

# Save password only in session
if st.session_state.remember_pass_session and email_pass:
    st.session_state.saved_pass = email_pass


# -------------------------------------------------------
# EMAIL DETAILS
# -------------------------------------------------------
st.header("2. Email Details")

subject = st.text_input("Email Subject")

st.markdown("### Email Body (Rich Text Editor)")
editor_output = st_quill(
    value=st.session_state.body_html if not st.session_state.quill_initialized else "",
    html=True,
    placeholder="Write your email here...",
    key="MAIN_EDITOR"
)
st.session_state.quill_initialized = True

if editor_output and editor_output != st.session_state.body_html:
    st.session_state.body_html = editor_output


# -------------------------------------------------------
# EXCEL UPLOAD
# -------------------------------------------------------
st.header("3. Upload Recipient Excel File")
uploaded_excel = st.file_uploader("Upload .xlsx file", type=["xlsx"])

df = None
detected_fields = {}

if uploaded_excel:
    try:
        df = pd.read_excel(uploaded_excel)
        detected_fields = detect_columns(df)
        st.success(f"Excel uploaded — {len(df)} rows found.")
        st.info(f"Detected fields: {detected_fields}")
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")


# -------------------------------------------------------
# PREVIEW
# -------------------------------------------------------
st.header("4. Preview")

if st.button("Show Preview"):
    if not st.session_state.body_html.strip():
        st.error("Email body is empty.")
    else:
        preview = st.session_state.body_html

        if df is not None and len(df) > 0:
            first = df.iloc[0]
            for field, col in detected_fields.items():
                preview = preview.replace(
                    f"{{{field}}}",
                    "" if pd.isna(first[col]) else str(first[col])
                )

        st.markdown("### Email Preview")
        st.markdown(preview, unsafe_allow_html=True)


# -------------------------------------------------------
# ATTACHMENTS
# -------------------------------------------------------
st.header("5. Attachments (optional)")
uploaded_files = st.file_uploader(
    "Upload attachments",
    type=["pdf", "jpeg", "jpg", "png"],
    accept_multiple_files=True
)


# -------------------------------------------------------
# SEND EMAILS
# -------------------------------------------------------
st.header("6. Send Emails")

if st.button("🚀 Send Now"):
    if df is None:
        st.error("Please upload an Excel file.")
        st.stop()

    if "email" not in detected_fields:
        st.error("No email column detected.")
        st.stop()

    if not email_user or not email_pass:
        st.error("Email + App Password required.")
        st.stop()

    if not subject.strip():
        st.error("Subject cannot be empty.")
        st.stop()

    if not st.session_state.body_html.strip():
        st.error("Email body cannot be empty.")
        st.stop()

    # Save attachments temporarily
    temp_paths = []
    for file in uploaded_files:
        ext = os.path.splitext(file.name)[1]
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
        tmp.write(file.read())
        tmp.close()
        temp_paths.append(tmp.name)

    # Connect SMTP
    try:
        yag = yagmail.SMTP(email_user, email_pass)
    except Exception as e:
        st.error(f"SMTP Login Failed: {e}")
        st.stop()

    logs = []
    total = len(df)
    progress = st.progress(0)
    count = 0

    email_col = detected_fields["email"]

    for idx, row in df.iterrows():
        body = st.session_state.body_html

        for field, col in detected_fields.items():
            body = body.replace(
                f"{{{field}}}",
                "" if pd.isna(row[col]) else str(row[col])
            )

        # Split multiple emails
        raw_emails = re.split(r"[\/,; ]+", str(row[email_col]))
        emails = [e for e in raw_emails if "@" in e]

        if not emails:
            logs.append([idx, "", "NO_EMAIL", "SKIPPED", datetime.utcnow()])
            count += 1
            progress.progress(count / total)
            continue

        for email_addr in emails:
            try:
                yag.send(
                    to=email_addr,
                    subject=subject,
                    contents=body,
                    attachments=temp_paths
                )
                logs.append([idx, email_addr, "SENT", "OK", datetime.utcnow()])
            except Exception as err:
                logs.append([idx, email_addr, "FAILED", str(err), datetime.utcnow()])

        count += 1
        progress.progress(count / total)

    # EXPORT CLEAN EXCEL LOG
    excel_path, excel_name = export_logs_excel(logs)

    st.success("All emails processed!")
    with open(excel_path, "rb") as f:
        st.download_button("📥 Download Logs (Excel)", f, file_name=excel_name)

    # Cleanup temp files
    for p in temp_paths:
        try:
            os.unlink(p)
        except:
            pass
