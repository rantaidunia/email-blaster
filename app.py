#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import yagmail
import re
import tempfile
import os
from io import StringIO
from streamlit_quill import st_quill
from datetime import datetime

# ---------------------------
# CONFIG
# ---------------------------
st.set_page_config(page_title="Email Blaster", layout="wide")

st.title("📧 Email Blasters")

st.markdown(
    """
Upload an Excel (.xlsx) with an **email** column.  
Use placeholders like `{name}`, `{position}`, `{company}`, etc — fields will be detected automatically.
"""
)

# ---------------------------
# FIELD DETECTION
# ---------------------------
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

        # Strict match for email
        if field == "email":
            for norm_col, real_col in normalized_cols.items():
                if norm_col == "email":
                    detected[field] = real_col
                    break
            if field in detected:
                continue

        # Flexible match for others
        for alias in aliases:
            norm_alias = normalize(alias)
            for norm_col, real_col in normalized_cols.items():
                if norm_alias in norm_col or norm_col in norm_alias:
                    detected[field] = real_col
                    break
            if field in detected:
                break

    return detected



# ---------------------------
# 1) Email login
# ---------------------------
st.header("1. Email Account Login")
email_user = st.text_input("Your Email Address", placeholder="example@gmail.com")
email_pass = st.text_input("App Password (NOT your normal password)", type="password")
st.info("For Gmail: create an App Password at https://myaccount.google.com/apppasswords")

# ---------------------------
# 2) Upload Excel
# ---------------------------
st.header("2. Upload Recipient Excel File")
uploaded_excel = st.file_uploader("Upload .xlsx file", type=["xlsx"])

df = None
detected_fields = {}

if uploaded_excel:
    try:
        df = pd.read_excel(uploaded_excel)
        detected_fields = detect_columns(df)

        st.success(f"Excel uploaded successfully — {len(df)} rows loaded.")

        # Show detected fields
        st.info(f"Detected fields: {detected_fields}")

        if "email" not in detected_fields:
            st.error("No valid email column detected. Please check your Excel headers.")
    except Exception as e:
        st.error(f"Failed to load Excel: {e}")

# ---------------------------
# 3) Email Details + Quill editor (FINAL FIX)
# ---------------------------
st.header("3. Email Details")

subject = st.text_input("Email Subject")

st.markdown("### Email Body (Rich Text Editor)")

# Initialize only once the FIRST time Quill loads
if "body_html" not in st.session_state:
    st.session_state.body_html = ""
if "quill_loaded" not in st.session_state:
    st.session_state.quill_loaded = False

# Only supply initial value ONCE — never again
quill_initial_value = None
if not st.session_state.quill_loaded:
    quill_initial_value = st.session_state.body_html

# Render Quill
email_body_html = st_quill(
    value=quill_initial_value,
    placeholder="Write your email here... Use {name}, {company}, etc.",
    html=True,
    key="quill_body_editor"
)

# Mark as loaded so next reruns do NOT reset the value
if not st.session_state.quill_loaded:
    st.session_state.quill_loaded = True

# Store updated content
if email_body_html and email_body_html != st.session_state.body_html:
    st.session_state.body_html = email_body_html


# ---------------------------
# 4) Preview
# ---------------------------
st.header("4. Preview")

if st.button("Show Preview"):
    if not st.session_state["body_html"].strip():
        st.error("Please write an email body first.")
    else:
        preview = st.session_state["body_html"]

        if df is not None and len(df) > 0:
            row = df.iloc[0]

            for field, excel_col in detected_fields.items():
                placeholder = f"{{{field}}}"
                value = "" if pd.isna(row[excel_col]) else str(row[excel_col])
                preview = preview.replace(placeholder, value)

        st.markdown("### Email Preview")
        st.markdown(preview, unsafe_allow_html=True)

# ---------------------------
# 5) Attachments
# ---------------------------
st.header("5. Attachments (optional)")
uploaded_files = st.file_uploader(
    "Upload attachments",
    type=["pdf", "jpeg", "jpg", "png"],
    accept_multiple_files=True
)

# ---------------------------
# 6) Send Emails
# ---------------------------
st.header("6. Send Emails")

if st.button("🚀 Send Now"):
    if df is None:
        st.error("Please upload an Excel file.")
        st.stop()

    if "email" not in detected_fields:
        st.error("No valid email field found in Excel.")
        st.stop()

    if not email_user or not email_pass:
        st.error("Please enter your email and app password.")
        st.stop()

    if not subject.strip():
        st.error("Please enter an email subject.")
        st.stop()

    if not st.session_state["body_html"].strip():
        st.error("Please write an email body.")
        st.stop()

    # Save attachments
    temp_paths = []
    try:
        for f in uploaded_files:
            suffix = os.path.splitext(f.name)[1]
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            tmp.write(f.read())
            tmp.close()
            temp_paths.append(tmp.name)
    except Exception as e:
        st.error(f"Attachment error: {e}")
        for p in temp_paths:
            try: os.unlink(p)
            except: pass
        st.stop()

    # Connect SMTP
    st.info("Connecting to SMTP...")
    try:
        yag = yagmail.SMTP(email_user, email_pass)
    except Exception as e:
        st.error(f"SMTP Login Failed: {e}")
        st.stop()

    # Logs
    logs = []
    total = len(df)
    progress = st.progress(0)
    count = 0

    email_col = detected_fields["email"]

    # Sending loop
    for idx, row in df.iterrows():
        body = st.session_state["body_html"]

        # Replace placeholders using detected fields
        for field, excel_col in detected_fields.items():
            placeholder = f"{{{field}}}"
            value = "" if pd.isna(row[excel_col]) else str(row[excel_col])
            body = body.replace(placeholder, value)

        # Parse emails inside cell
        raw = str(row[email_col])
        emails = re.split(r"[\/,; ]+", raw)
        emails = [e for e in emails if "@" in e]

        if not emails:
            logs.append([idx, raw, "NO_VALID_EMAIL", "SKIPPED", datetime.utcnow().isoformat()])
            count += 1
            progress.progress(count / total)
            continue

        # Send to each email individually
        for e in emails:
            try:
                yag.send(
                    to=e,
                    subject=subject,
                    contents=body,
                    attachments=temp_paths
                )
                logs.append([idx, e, "SENT", "OK", datetime.utcnow().isoformat()])
            except Exception as error:
                logs.append([idx, e, "FAILED", str(error), datetime.utcnow().isoformat()])

        count += 1
        progress.progress(count / total)

    # Save logs
    log_df = pd.DataFrame(logs, columns=["Row", "Email", "Status", "Details", "Timestamp"])
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"logs_{ts}.csv"

    csv_buf = StringIO()
    log_df.to_csv(csv_buf, index=False)

    st.success("All emails processed!")
    st.download_button("📥 Download Logs", csv_buf.getvalue(), filename, "text/csv")

    # Cleanup
    for p in temp_paths:
        try: os.unlink(p)
        except: pass




