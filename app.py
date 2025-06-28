import streamlit as st
import pandas as pd
import smtplib
import os
import json
import mimetypes
from email.message import EmailMessage

st.set_page_config(page_title="🕊️ Pegion – Smart Email Sender", layout="centered")

CREDENTIALS_FILE = "secrets.json"

def load_credentials():
    if os.path.exists(CREDENTIALS_FILE):
        with open(CREDENTIALS_FILE, "r") as f:
            data = json.load(f)
            return {"email": data.get("email", ""), "password": data.get("password", "")}
    return {"email": "", "password": ""}

def save_credentials(email, password):
    with open(CREDENTIALS_FILE, "w") as f:
        json.dump({"email": email, "password": password}, f)

def send_email(sender_email, sender_password, row, default_message, default_link, default_attachment_path):
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = row["Email"]
    msg["Subject"] = row.get("Subject", "No Subject")

    # 🔁 Fallback for missing or empty Message Template
    template = row.get("Message Template")
    if pd.isna(template) or not str(template).strip():
        template = default_message

    try:
        message = template.format(**row)
    except Exception:
        return False, "Template format error"

    # 🔗 Optional link support
    link = row.get("Link")
    if pd.isna(link) or not str(link).strip():
        link = default_link.strip()
    if link:
        message += f'<br><br><a href="{link}" target="_blank">🔗 Open Link</a>'

    msg.set_content(message)
    msg.add_alternative(f"<html><body>{message.replace(chr(10), '<br>')}</body></html>", subtype='html')

    # 📎 Attachments
    attachment_path = None
    if pd.notna(row.get("Attachment")) and str(row["Attachment"]).strip():
        attachment_path = os.path.join(os.getcwd(), str(row["Attachment"]).strip())
    elif default_attachment_path:
        attachment_path = default_attachment_path

    if attachment_path and os.path.isfile(attachment_path):
        try:
            with open(attachment_path, 'rb') as f:
                mime_type, _ = mimetypes.guess_type(attachment_path)
                mime_type = mime_type or 'application/octet-stream'
                maintype, subtype = mime_type.split('/', 1)
                msg.add_attachment(f.read(), maintype=maintype, subtype=subtype,
                                   filename=os.path.basename(attachment_path))
        except Exception as e:
            return False, f"Attachment error: {e}"

    # 📤 Send email
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
        return True, ""
    except Exception as e:
        return False, str(e)

# === UI Starts ===
st.title("🕊️ Pegion – Smart Email Sender")
st.markdown("Deliver personalized emails from your Excel sheet like a true digital postman! 🧾")

# 🔐 Credentials
creds = load_credentials()
with st.expander("🔐 Email Credentials", expanded=(not creds["email"] or not creds["password"])):
    with st.form("credentials_form"):
        input_email = st.text_input("Your Gmail Address", value=creds["email"])
        input_password = st.text_input("App Password", type="password", value=creds["password"])
        save = st.form_submit_button("✅ Save & Continue")
        if save:
            save_credentials(input_email, input_password)
            st.success("Credentials saved! Please reload the app.")

if st.button("❌ Clear Saved Credentials"):
    if os.path.exists(CREDENTIALS_FILE):
        os.remove(CREDENTIALS_FILE)
        st.success("Credentials cleared. Please reload the app.")

creds = load_credentials()
sender_email = creds["email"]
app_password = creds["password"]

if sender_email and app_password:
    st.markdown("### 📁 Upload Excel File")
    uploaded_file = st.file_uploader("Choose an Excel (.xlsx) file", type=["xlsx"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip()

            if not {'Email', 'Subject'}.issubset(df.columns):
                st.error("Excel must contain at least: Email, Subject")
            else:
                st.success("Excel loaded successfully!")
                st.dataframe(df)

                # ✅ Ensure fallback column exists
                if 'Message Template' not in df.columns:
                    df['Message Template'] = None

                st.markdown("### ✏️ Default Inputs (Optional)")
                default_message = st.text_area("Default Message Template", value="Hello <b>{Name}</b>, welcome!", height=150)
                default_link = st.text_input("Default Link", placeholder="https://example.com")

                default_attachment = st.file_uploader("Upload Default Attachment (optional)", type=None)
                default_attachment_path = None
                if default_attachment:
                    default_attachment_path = os.path.join(os.getcwd(), default_attachment.name)
                    with open(default_attachment_path, "wb") as f:
                        f.write(default_attachment.read())

                if st.button("🚀 Send Emails"):
                    results = []
                    progress = st.progress(0, text="Starting to send emails...")

                    for i, row in df.iterrows():
                        success, error = send_email(sender_email, app_password, row, default_message, default_link, default_attachment_path)
                        results.append({
                            "Name": row.get("Name", ""),
                            "Email": row["Email"],
                            "Status": "✅ Sent" if success else "❌ Failed",
                            "Error": error
                        })
                        progress.progress((i + 1) / len(df), text=f"Sent {i+1} of {len(df)}")

                    progress.empty()
                    result_df = pd.DataFrame(results)
                    st.subheader("📋 Email Status")
                    st.dataframe(result_df)

                    failed_df = result_df[result_df["Status"] == "❌ Failed"]
                    if not failed_df.empty:
                        if st.button("🔁 Retry Failed Emails"):
                            st.info(f"Retrying {len(failed_df)} emails...")
                            for i, row in failed_df.iterrows():
                                original_row = df[df["Email"] == row["Email"]].iloc[0]
                                success, error = send_email(sender_email, app_password, original_row, default_message, default_link, default_attachment_path)
                                result_df.loc[row.name, "Status"] = "✅ Retried" if success else "❌ Failed"
                                result_df.loc[row.name, "Error"] = error

                            st.success("✅ Retry complete!")
                            st.dataframe(result_df)

        except Exception as e:
            st.error(f"❌ Error loading file: {e}")
else:
    st.info("🔐 Please enter and save your credentials to proceed.")

st.markdown("---")
st.markdown("Crafted with ❤️ by Bharat | Powered by Pegion 🕊️", unsafe_allow_html=True)
