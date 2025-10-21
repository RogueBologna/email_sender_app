# Email Sender App

Modern desktop email client built with Python and Tkinter. Send emails via Gmail with attachments, templates, and dark/light mode toggle.

## Features
- ✅ GUI with recipient lists, subjects, and rich text body
- ✅ Attachment support
- ✅ Template/history saving (JSON-based)
- ✅ Credential auto-save (optional)
- ✅ Error logging
- ✅ Cross-platform (Windows/Mac/Linux)

## Demo
<image-card alt="Screenshot" src="screenshots/app-screenshot.png" ></image-card>  <!-- Add a screenshot here later -->

## Installation
1. Download the executable from Releases or run with Python.
2. `pip install -r requirements.txt`
3. Run: `python main.py`

## Setup for Gmail
- Enable 2FA on your Google account.
- Generate an [App Password](https://myaccount.google.com/apppasswords).

## Build Executable
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "EmailSender" main.py