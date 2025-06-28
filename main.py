import pandas as pd
import smtplib
from email.message import EmailMessage
import os

# === CONFIGURATION ===
EMAIL_ADDRESS = 'bharathmallipuram6@gmail.com'
EMAIL_PASSWORD = 'yfhg gixp cpyu fbix'
EXCEL_FILE = 'email_list.xlsx'            # Excel File

# === LOAD EXCEL SAFELY ===
try:
    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip()  # Clean up column names
except Exception as e:
    print(f"❌ Failed to read Excel file: {e}")
    exit(1)

# === CHECK REQUIRED COLUMNS ===
required_columns = ['Email', 'Subject', 'Message Template']
missing_columns = [col for col in required_columns if col not in df.columns]

if missing_columns:
    print(f"❌ Missing columns in Excel: {missing_columns}")
    exit(1)

# === START EMAIL LOOP ===
for index, row in df.iterrows():
    email_to = row.get('Email')
    subject = row.get('Subject')
    attachment_file = row.get('Attachment') if 'Attachment' in row else None

    # === Build Personalized Message ===
    try:
        message_template = row.get('Message Template', '').strip()
        if not message_template:
            raise ValueError("Message template is empty")
        message = message_template.format(**row)
    except Exception as e:
        print(f"⚠️ Error formatting message for {email_to}: {e}")
        continue

    # === Compose Email ===
    msg = EmailMessage()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = email_to
    msg['Subject'] = subject or "No Subject"

    msg.set_content(message)  # Fallback plain text
    html_body = f"""\
    <html>
      <body>
        <p>{message.replace('\n', '<br>')}</p>
      </body>
    </html>
    """
    msg.add_alternative(html_body, subtype='html')

    # === Attach File If Exists ===
    if pd.notna(attachment_file) and str(attachment_file).strip():
        attachment_path = str(attachment_file).strip()
        if not os.path.isabs(attachment_path):
            attachment_path = os.path.join(os.getcwd(), attachment_path)
        if os.path.isfile(attachment_path):
            try:
                with open(attachment_path, 'rb') as f:
                    file_data = f.read()
                    file_name = os.path.basename(attachment_path)
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
                print(f"📎 Attached: {file_name}")
            except Exception as e:
                print(f"⚠️ Failed to attach file for {email_to}: {e}")
        else:
            print(f"⚠️ Attachment not found for {email_to}: {attachment_path}")

    # === SEND EMAIL ===
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print(f"✅ Email sent to {email_to}")
    except Exception as e:
        print(f"❌ Failed to send email to {email_to}: {e}")
