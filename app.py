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

def log_to_excel(recipient, status, details):
    log_filename = "email_logs.xlsx"

    # if not exist, create workbook + headers
    if not os.path.exists(log_filename):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Logs"

        headers = ["Timestamp", "Recipient", "Status", "Details"]
        ws.append(headers)

        # style headers
        for col in range(1, len(headers) + 1):
            ws.cell(row=1, column=col).font = Font(bold=True)

        wb.save(log_filename)

    # open existing log
    wb = openpyxl.load_workbook(log_filename)
    ws = wb.active

    # append new row
    ws.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        recipient,
        status,
        details
    ])

    # auto column width
    for col in ws.columns:
        max_len = 0
        column = col[0].column
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(column)].width = max_len + 2

    # add borders
    thin = Border(left=Side(style="thin"),
                  right=Side(style="thin"),
                  top=Side(style="thin"),
                  bottom=Side(style="thin"))

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin

    wb.save(log_filename)

# -------------------------------------------------------
# CONFIG
# -------------------------------------------------------
st.set_page_config(page_title="Email Blaster", layout="wide")

st.title("📧 Email Blaster")

st.markdown("""
Upload an Excel (.xlsx) with an **email** column.  
Use placeholders like `{name}`, `{position}`, `{company}`, etc.
""")

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
            for norm_col, real_col in normalized_cols.items():
                if norm_col == "email":
                    detected[field] = real_col
                    break
            if field in detected:
                continue

        for alias in aliases:
            norm_alias = normalize(alias)
            for norm_col, real_col in normalized_cols.items():
                if norm_alias in norm_col or norm_col in norm_alias:
                    detected[field] = real_col
                    break
            if field in detected:
                break

    return detected


# -------------------------------------------------------
# QUILL PREP
# ensure session state keys exist BEFORE any reruns that may come from widgets below
# -------------------------------------------------------
if "body_html" not in st.session_state:
    st.session_state.body_html = ""  # persisted HTML content

if "quill_initialized" not in st.session_state:
    st.session_state.quill_initialized = False  # whether we already passed initial value


# -------------------------------------------------------
# EMAIL LOGIN (keep this early)
# -------------------------------------------------------
st.header("1. Email Account Login")
email_user = st.text_input("Your Email Address", placeholder="example@gmail.com")
email_pass = st.text_input("App Password (NOT your normal password)", type="password")
st.info("For Gmail: create an App Password at https://myaccount.google.com/apppasswords")


# -------------------------------------------------------
# 3) Email Details — render Quill BEFORE file uploader
# This is IMPORTANT: Quill must mount before the uploader to avoid re-init races.
# -------------------------------------------------------
st.header("2. Email Details")

subject = st.text_input("Email Subject")

st.markdown("### Email Body (Rich Text Editor)")

# --- Critical pattern:
# 1) On first load, call st_quill with value=st.session_state.body_html
# 2) On subsequent reruns, call st_quill WITHOUT the value parameter
#    (calling with value=None or value="" after init breaks the editor)
editor_output = None
if not st.session_state.quill_initialized:
    # first-time mounting: provide the saved HTML (may be empty)
    editor_output = st_quill(
        value=st.session_state.body_html,
        html=True,
        placeholder="Write your email here... Use {name}, {company}, etc.",
        key="MAIN_EDITOR"
    )
    st.session_state.quill_initialized = True
else:
    # subsequent renders: do NOT pass 'value' argument
    editor_output = st_quill(
        html=True,
        placeholder="Write your email here... Use {name}, {company}, etc.",
        key="MAIN_EDITOR"
    )

# store only when editor gives us content (avoid overwriting with empty)
if editor_output and editor_output != st.session_state.body_html:
    st.session_state.body_html = editor_output


# -------------------------------------------------------
# EXCEL UPLOAD (now safe to place after Quill)
# -------------------------------------------------------
st.header("3. Upload Recipient Excel File")
uploaded_excel = st.file_uploader("Upload .xlsx file", type=["xlsx"])

df = None
detected_fields = {}

if uploaded_excel:
    try:
        df = pd.read_excel(uploaded_excel)
        detected_fields = detect_columns(df)

        st.success(f"Excel uploaded successfully — {len(df)} rows loaded.")
        st.info(f"Detected fields: {detected_fields}")

        if "email" not in detected_fields:
            st.error("No valid email field found. Fix headers in your Excel.")
    except Exception as e:
        st.error(f"Failed to process Excel: {e}")


# -------------------------------------------------------
# PREVIEW (uses body_html which is stable)
# -------------------------------------------------------
st.header("4. Preview")

if st.button("Show Preview"):
    if not st.session_state.body_html.strip():
        st.error("Please write the email body first.")
    else:
        preview = st.session_state.body_html

        if df is not None and len(df) > 0:
            row = df.iloc[0]

            for field, col in detected_fields.items():
                placeholder = f"{{{field}}}"
                value = "" if pd.isna(row[col]) else str(row[col])
                preview = preview.replace(placeholder, value)

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
        st.error("No valid email column detected.")
        st.stop()

    if not email_user or not email_pass:
        st.error("Email + App password required.")
        st.stop()

    if not subject.strip():
        st.error("Subject required.")
        st.stop()

    if not st.session_state.body_html.strip():
        st.error("Email body is empty.")
        st.stop()

    # Save uploaded attachments
    temp_paths = []
    try:
        for file in uploaded_files:
            ext = os.path.splitext(file.name)[1]
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
            tmp.write(file.read())
            tmp.close()
            temp_paths.append(tmp.name)
    except Exception as e:
        st.error("Attachment error.")
        st.stop()

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
            body = body.replace(f"{{{field}}}", "" if pd.isna(row[col]) else str(row[col]))

        raw_emails = re.split(r"[\/,; ]+", str(row[email_col]))
        emails = [e for e in raw_emails if "@" in e]

        if not emails:
            logs.append([idx, raw_emails, "NO_EMAIL", "SKIPPED", datetime.utcnow()])
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
        log_to_excel(email_addr, "SENT", "OK")   # <--- NEW Excel log
    except Exception as err:
        logs.append([idx, email_addr, "FAILED", str(err), datetime.utcnow()])
        log_to_excel(email_addr, "FAILED", str(err))   # <--- NEW Excel log

        count += 1
        progress.progress(count / total)

    # Save logs
    log_df = pd.DataFrame(logs, columns=["Row", "Email", "Status", "Details", "Timestamp"])
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"logs_{ts}.csv"

    st.success("All emails processed!")
    st.download_button("📥 Download Logs", log_df.to_csv(index=False), filename, "text/csv")

    # Cleanup temp
    for p in temp_paths:
        try: os.unlink(p)
        except: pass


