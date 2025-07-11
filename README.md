# 🕊️ Pegion – Smart Email Sender

**Pegion** is a lightweight, intelligent email sender built using Streamlit that allows you to send **personalized, HTML-formatted emails** to thousands of people in just a few clicks — with optional attachments, fallback values, retry support, and smart Google Maps integration.

> 📫 Ideal for events, convocations, workshops, and institutional campaigns.

---

## ✨ Features

- ✅ Upload Excel with dynamic fields like `{Name}`, `{Email}`, `{Subject}`
- ✅ Personalized message templating per recipient
- ✅ Default fallback for message, link, and attachment
- ✅ Retry failed emails with one click
- ✅ Dynamic **Google Maps Directions** link (based on recipient’s live location)
- ✅ HTML-rich email body with preview support
- ✅ One-time Gmail App Password login (secure)
- ✅ Clean UI, no manual refreshes, no credential re-entry

---

## 📂 Folder Structure
### 2. Install dependencies
```bash
Copy
Edit
pip install streamlit pandas openpyxl
```
### 3. Run the app
```bash
Copy
Edit
streamlit run Pegion.py
```

## 🔐 Gmail Setup
Go to Google App Passwords

Generate a 16-digit App Password

Paste it into the app's credentials section

##⚠️ Do NOT use your regular Gmail password!

## 🗺️ Dynamic Maps Integration
Pegion auto-injects a clickable “📍 Get Directions” button in each email, opening:

```ruby
Copy
Edit
https://www.google.com/maps/dir/?api=1&destination=Your+Venue+Address
```
📌 Users just click it → their Maps opens with live directions from their current location.

🛡️ Security & Notes
Your credentials are stored locally in secrets.json

Delete them anytime via “Clear Credentials” button

This app uses secure Gmail SMTP with SSL (port 465)

## ❤️ Credits
Built by Bharat to automate and simplify email campaigns for convocation events and beyond.

## 📬 Sample Email Preview
```html
Copy
Edit
Hello <b>{Name}</b>,<br><br>
You're invited to the Grand Convocation.<br>
<a href="https://your-event.com">🔗 View Full Details</a><br>
<a href="https://maps.google.com/...">📍 Get Directions to Venue</a>
```
## 📌 To-Do / Future Features
📅 Schedule emails for future

📊 Analytics dashboard

🌐 Multi-language email support

🧠 AI assistant for message generation

Feel free to fork, use, and scale it for your institution or startup.
Let Pegion do the flying! 🕊️

```yaml
Copy
Edit

---

Let me know if you want this included as a downloadable `README.md` inside the app, Bharat!
```







Ask ChatGPT
