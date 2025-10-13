import streamlit as st
import pandas as pd
import base64
import time
import random
import re
import json
import pytz
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="centered", page_icon="📧")
st.title("📧 Gmail Mail Merge")

# ========================================
# File Upload
# ========================================
uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])
if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    st.success(f"✅ Loaded {len(df)} records from file.")
else:
    st.info("Please upload your mail merge CSV to begin.")

# ========================================
# Gmail OAuth Setup
# ========================================
CLIENT_SECRETS_FILE = "client_secret.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

def authenticate_gmail():
    flow = Flow.from_client_secrets_file(CLIENT_SECRETS_FILE, scopes=SCOPES, redirect_uri="urn:ietf:wg:oauth:2.0:oob")
    auth_url, _ = flow.authorization_url(prompt="consent")
    st.markdown(f"[Authorize Gmail here]({auth_url})")
    auth_code = st.text_input("Enter the authorization code:")
    if auth_code:
        flow.fetch_token(code=auth_code)
        creds = flow.credentials
        st.success("✅ Gmail authenticated successfully!")
        return creds
    return None

# ========================================
# Gmail Message Builder
# ========================================
def create_message(sender, to, subject, message_text, thread_id=None, reply_to=None):
    message = MIMEText(message_text, "html")
    message["to"] = to
    message["from"] = sender
    message["subject"] = subject
    if reply_to:
        message["In-Reply-To"] = reply_to
        message["References"] = reply_to
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    body = {"raw": raw_message}
    if thread_id:
        body["threadId"] = thread_id
    return body

def send_message(service, user_id, message_body):
    try:
        sent_message = service.users().messages().send(userId=user_id, body=message_body).execute()
        return sent_message
    except Exception as e:
        st.error(f"Error sending message: {e}")
        return None

# ========================================
# Label & Timing Options
# ========================================
st.header("🏷️ Label & Timing Options")

label_name = st.text_input("Gmail label to apply (new emails only)", "Mail Merge Sent")
delay = st.slider("Delay between emails (seconds)", 5, 60, 30)

# ========================================
# ✅ Enhanced Estimated Completion Time (±10%) — Local Time Fixed
# ========================================
if uploaded_file is not None:
    try:
        total_contacts = len(df)
        avg_delay = delay

        # Randomization bounds
        min_delay = avg_delay * 0.9
        max_delay = avg_delay * 1.1

        # Total time bounds
        min_total_seconds = total_contacts * min_delay
        max_total_seconds = total_contacts * max_delay
        avg_total_seconds = total_contacts * avg_delay

        avg_total_minutes = avg_total_seconds / 60
        avg_total_hours = avg_total_minutes / 60

        # ✅ Local time (IST)
        local_tz = pytz.timezone("Asia/Kolkata")
        now_local = datetime.now(local_tz)

        eta_min = now_local + timedelta(seconds=min_total_seconds)
        eta_max = now_local + timedelta(seconds=max_total_seconds)

        eta_min_str = eta_min.strftime("%I:%M %p")
        eta_max_str = eta_max.strftime("%I:%M %p")

        if avg_total_hours >= 1:
            duration_text = f"⏳ {avg_total_hours:.2f} hr (±10%)"
        else:
            duration_text = f"⏳ {avg_total_minutes:.1f} min (±10%)"

        st.caption(
            f"📋 Total: {total_contacts} | {duration_text} | 🕒 ETA: **{eta_min_str} – {eta_max_str} (IST)**"
        )
    except Exception:
        st.caption("⚠️ ETA unavailable — check timezone or input data.")

# ========================================
# Sending Mode Options
# ========================================
send_mode = st.radio(
    "Choose sending mode",
    ["New Email", "Follow-up (Reply)", "Save as Draft"],
    horizontal=False,
)

# ========================================
# Send / Draft Execution
# ========================================
if st.button("🚀 Send Emails / Save Drafts"):
    if uploaded_file is None:
        st.error("Please upload your CSV file first.")
    else:
        creds = authenticate_gmail()
        if creds:
            service = build("gmail", "v1", credentials=creds)
            sender = "me"

            for index, row in df.iterrows():
                to = row.get("To")
                subject = row.get("Subject", "")
                message_text = row.get("Body", "")
                thread_id = row.get("ThreadId", None)
                reply_to = row.get("MessageId", None)

                if send_mode == "Follow-up (Reply)" and not thread_id:
                    st.warning(f"⚠️ Row {index + 1}: Missing thread/message ID, skipped.")
                    continue

                message_body = create_message(sender, to, subject, message_text, thread_id, reply_to)

                if send_mode == "Save as Draft":
                    service.users().drafts().create(userId="me", body={"message": message_body}).execute()
                    st.info(f"💾 Draft saved for: {to}")
                else:
                    result = send_message(service, "me", message_body)
                    if result:
                        st.success(f"✅ Sent to: {to}")

                randomized_delay = random.uniform(delay * 0.9, delay * 1.1)
                time.sleep(randomized_delay)

            st.balloons()
            st.success("🎉 All emails processed successfully!")
