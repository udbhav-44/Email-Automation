# Email-Automation

A Python-based automation tool for sending personalized emails to multiple recipients using data from Google Sheets.

## Features
- Reads recipient data (professor names, email addresses, and personalized writeups) from Google Sheets
- Sends personalized emails using SMTP
- Tracks sent emails to avoid duplicates
- Supports environment variables for secure credential management

## Setup
1. Create a `.env` file based on `.env.example` with your email credentials:
```
SMTP_SERVER=your_smtp_server
SMTP_PORT=your_smtp_port
SENDER_EMAIL=your_email
EMAIL_PASSWORD=your_password
CREDENTIALS_FILE=path_to_google_credentials.json
```

2. Set up Google Sheets API:
   - Create a service account and download credentials JSON
   - Share your Google Sheet with the service account email
   - Update the credentials file path in your `.env`

## Usage
1. Prepare your Google Sheet with columns: "Professor", "Mail Id", and "Writeup"
2. Run `python mailing.py` to send emails

