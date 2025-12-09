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
# EMAIL LOGIN
# -------------------------------------------------------
st.header("1. Email Account Login")
email_user = st.text_input("Your Email Address", placeholder="example@gmail.com")
email_pass = st.text_input("App Password (NOT your normal password)", type="password")
st.info("For Gmail: create an App Password at https://myaccount.google.com/apppasswords")

# -------------------------------------------------------
# EXCEL UPLOAD
# -------------------------------------------------------
st.header("2. Upload Recipient Excel File")
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
# QUILL EDITOR — FULLY FIXED
# -------------------------------------------------------
st.header("3. Email Details")

subject = st.text_input("Email Subject")

st.markdown("### Email Body (Rich Text Editor)")

# Initialize session state once
if "body_html" not in st.session_state:
    st.session_state.body_html = ""
if "quill_initialized" not in st.session_state:
    st.session_state.quill_initialized = False

# Provide initial value ONLY ONCE
initial_quill_value = None
if not st.session_state.quill_initialized:
    initial_quill_value = st.session_state.body_html

# Render Quill
body_html = st_quill(
    value=initial_quill_value,
    placeholder="Write your email here... Use {name}, {company}, etc.",
    html=True,
    key="email_quill_editor"
)

# Mark initialized
if not st.session_state.quill_initialized:
    st.session_state.quill_initialized = True

# Update stored HTML only if changed
if body_html and body_html != st.session_state.body_html:
    st.session_state.body_html = body_html


# -------------------------------------------------------
# PREVIEW
# -------------------------------------------------------
st.header("4. Preview")

if st.button("Show Preview"):
    if not st.session_state.body_html.strip():
        st.error("Please write the email body first.")
    else:
        preview = st.session_state.body_html

        if df is not None and len(df) > 0:
            row = df.iloc[0]  # First row only

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
        st.error("Email and app password required.")
        st.stop()

    if not subject.strip():
        st.error("Subject required.")
        st.stop()

    if not st.session_state.body_html.strip():
        st.error("Email body is empty.")
        st.stop()

    temp_paths = []
    try:
        for file in uploaded_files:
            ext = os.path.splitext(file.name)[1]
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
            tmp.write(file.read())
            tmp.close()
            temp_paths.append(tmp.name)
    except:
        st.error("Attachment error.")
        st.stop()

    # SMTP Connect
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

    for index, row in df.iterrows():
        body = st.session_state.body_html

        for field, col in detected_fields.items():
            val = "" if pd.isna(row[col]) else str(row[col])
            body = body.replace(f"{{{field}}}", val)

        raw_emails = re.split(r"[\/,; ]+", str(row[email_col]))
        emails = [e for e in raw_emails if "@" in e]

        if not emails:
            logs.append([index, raw_emails, "NO_EMAIL", "SKIPPED", datetime.utcnow()])
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
                logs.append([index, email_addr, "SENT", "OK", datetime.utcnow()])
            except Exception as err:
                logs.append([index, email_addr, "FAILED", str(err), datetime.utcnow()])

        count += 1
        progress.progress(count / total)

    # Save logs
    log_df = pd.DataFrame(logs, columns=["Row", "Email", "Status", "Details", "Timestamp"])
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"logs_{ts}.csv"

    csv_data = log_df.to_csv(index=False)

    st.success("All emails processed!")
    st.download_button("📥 Download Logs", csv_data, filename, "text/csv")

    # Clean temp
    for p in temp_paths:
        try: os.unlink(p)
        except: pass
