#!/usr/bin/env python
# coding: utf-8

import streamlit as st
import pandas as pd
import yagmail
import re
import tempfile
import os
from io import StringIO
from streamlit_quill import st_quill
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter

# -------------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------------
st.set_page_config(
    page_title="Email Blaster",
    layout="wide",
    page_icon="📧"
)

# -------------------------------------------------------
# CUSTOM CSS — MODERN UI
# -------------------------------------------------------
st.markdown(
    """
    <style>

    /* Global font & spacing */
    body, input, textarea {
        font-family: 'Inter', sans-serif !important;
    }

    /* Card-like container */
    .section-card {
        background: #ffffff;
        padding: 25px 30px;
        border-radius: 14px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
        margin-bottom: 25px;
        border: 1px solid #f2f2f2;
    }

    .header-title {
        font-size: 32px !important;
        font-weight: 700 !important;
        margin-bottom: 5px;
    }

    .subheader {
        font-size: 18px;
        font-weight: 600;
        margin-bottom: 15px;
        color:#444;
    }

    /* Fix streamlit's default wide spacing */
    .block-container {
        padding-top: 1rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }

    /* Better buttons */
    .stButton>button {
        border-radius: 10px;
        padding: 10px 18px;
        font-size: 16px;
        border: 0px;
        background: #4b8df8;
        color: white;
        font-weight: 600;
        transition: 0.2s;
    }
    .stButton>button:hover {
        background: #2f71e8;
        color: white;
    }

    /* Preview box */
    .preview-box {
        padding: 20px;
        border-radius: 12px;
        background: #fafafa;
        border: 1px solid #eee;
        margin-top: 15px;
    }

    /* Hide "Made with Streamlit" */
    #MainMenu {visibility:hidden;}
    footer {visibility:hidden;}

    </style>
    """,
    unsafe_allow_html=True
)

# -------------------------------------------------------
# HEADER
# -------------------------------------------------------
st.markdown("<div class='header-title'>📧 Email Blaster</div>", unsafe_allow_html=True)
st.write("A simple and powerful tool to send personalized bulk emails.")

# -------------------------------------------------------
# EXCEL LOGGING FUNCTION
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

    for col in range(1, len(headers) + 1):
        ws.cell(row=1, column=col).font = Font(bold=True)

    for row in logs:
        ws.append(row)

    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 2

    wb.save(filepath)
    return filepath, filename


# Remember states
if "saved_email" not in st.session_state: st.session_state.saved_email = None
if "saved_pass" not in st.session_state: st.session_state.saved_pass = None
if "remember_email" not in st.session_state: st.session_state.remember_email = False
if "remember_pass_session" not in st.session_state: st.session_state.remember_pass_session = False

query_params = st.experimental_get_query_params()
if "email" in query_params and st.session_state.saved_email is None:
    st.session_state.saved_email = query_params["email"][0]

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
        if field == "email":
            for norm, real in normalized_cols.items():
                if norm == "email":
                    detected[field] = real
                    break
            if field in detected:
                continue

        for alias in aliases:
            alias_norm = normalize(alias)
            for norm, real in normalized_cols.items():
                if alias_norm in norm or norm in alias_norm:
                    detected[field] = real
                    break

    return detected


# -------------------------------------------------------
# SECTION 1 — LOGIN
# -------------------------------------------------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("<div class='subheader'>1. Email Login</div>", unsafe_allow_html=True)

colA, colB = st.columns([1, 1])

with colA:
    st.session_state.remember_email = st.checkbox("Remember Email")
with colB:
    st.session_state.remember_pass_session = st.checkbox("Remember App Password (session only)")

email_user = st.text_input(
    "Your Email Address",
    value=st.session_state.saved_email if st.session_state.saved_email else "",
    placeholder="you@example.com"
)

email_pass = st.text_input(
    "App Password",
    type="password",
    value=st.session_state.saved_pass if st.session_state.remember_pass_session else ""
)

if st.session_state.remember_email and email_user:
    st.session_state.saved_email = email_user
    st.experimental_set_query_params(email=email_user)

if st.session_state.remember_pass_session and email_pass:
    st.session_state.saved_pass = email_pass

st.markdown("</div>", unsafe_allow_html=True)


# -------------------------------------------------------
# SECTION 2 — EMAIL CONTENT
# -------------------------------------------------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("<div class='subheader'>2. Email Content</div>", unsafe_allow_html=True)

subject = st.text_input("Email Subject")

if "body_html" not in st.session_state: st.session_state.body_html = ""
if "quill_initialized" not in st.session_state: st.session_state.quill_initialized = False

editor_output = st_quill(
    value=st.session_state.body_html if not st.session_state.quill_initialized else "",
    html=True,
    placeholder="Write your message here...",
    key="MAIN_EDITOR"
)

st.session_state.quill_initialized = True
if editor_output and editor_output != st.session_state.body_html:
    st.session_state.body_html = editor_output

st.markdown("</div>", unsafe_allow_html=True)


# -------------------------------------------------------
# SECTION 3 — EXCEL
# -------------------------------------------------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("<div class='subheader'>3. Upload Recipients</div>", unsafe_allow_html=True)

uploaded_excel = st.file_uploader("Upload .xlsx", type=["xlsx"])
df, detected_fields = None, {}

if uploaded_excel:
    df = pd.read_excel(uploaded_excel)
    detected_fields = detect_columns(df)
    st.success(f"Uploaded {len(df)} recipients.")
    st.info(f"Detected fields: {detected_fields}")

st.markdown("</div>", unsafe_allow_html=True)


# -------------------------------------------------------
# SECTION 4 — PREVIEW
# -------------------------------------------------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("<div class='subheader'>4. Preview</div>", unsafe_allow_html=True)

if st.button("Show Preview"):
    preview = st.session_state.body_html
    if df is not None and len(df) > 0:
        row = df.iloc[0]
        for field, col in detected_fields.items():
            preview = preview.replace(
                f"{{{field}}}",
                "" if pd.isna(row[col]) else str(row[col])
            )

    st.markdown("<div class='preview-box'>", unsafe_allow_html=True)
    st.markdown(preview, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)


# -------------------------------------------------------
# SECTION 5 — ATTACHMENTS
# -------------------------------------------------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("<div class='subheader'>5. Attachments</div>", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Upload attachments",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True
)
st.markdown("</div>", unsafe_allow_html=True)


# -------------------------------------------------------
# SECTION 6 — SEND EMAILS
# -------------------------------------------------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("<div class='subheader'>6. Send Emails</div>", unsafe_allow_html=True)

if st.button("🚀 Send Now"):
    if df is None:
        st.error("Upload an Excel file first.")
        st.stop()
    if "email" not in detected_fields:
        st.error("No email column detected.")
        st.stop()
    if not email_user or not email_pass:
        st.error("Email + App Password is required.")
        st.stop()
    if not subject.strip():
        st.error("Subject cannot be empty.")
        st.stop()
    if not st.session_state.body_html.strip():
        st.error("Email body cannot be empty.")
        st.stop()

    temp_paths = []
    for f in uploaded_files:
        ext = os.path.splitext(f.name)[1]
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
        tmp.write(f.read())
        tmp.close()
        temp_paths.append(tmp.name)

    try:
        yag = yagmail.SMTP(email_user, email_pass)
    except Exception as e:
        st.error(f"SMTP Login Failed → {e}")
        st.stop()

    logs = []
    total = len(df)
    progress = st.progress(0)

    email_col = detected_fields["email"]

    for idx, row in df.iterrows():
        body = st.session_state.body_html
        for field, col in detected_fields.items():
            body = body.replace(
                f"{{{field}}}",
                "" if pd.isna(row[col]) else str(row[col])
            )

        addresses = re.split(r"[\/,; ]+", str(row[email_col]))
        addresses = [a for a in addresses if "@" in a]

        for addr in addresses:
            try:
                yag.send(
                    to=addr,
                    subject=subject,
                    contents=body,
                    attachments=temp_paths
                )
                logs.append([idx, addr, "SENT", "OK", datetime.utcnow()])
            except Exception as e:
                logs.append([idx, addr, "FAILED", str(e), datetime.utcnow()])

        progress.progress((idx+1)/total)

    excel_path, excel_name = export_logs_excel(logs)
    st.success("Emails processed!")

    with open(excel_path, "rb") as f:
        st.download_button("📥 Download Log File", f, file_name=excel_name)

st.markdown("</div>", unsafe_allow_html=True)
