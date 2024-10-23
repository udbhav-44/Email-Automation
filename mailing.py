import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

smtp_server = "smtp.cc.iitk.ac.in"
smtp_port = 25
sender_email = "audbhav22@iitk.ac.in"
password = "Udbhav44"

# recipients = ["shubhangij22@iitk.ac.in", "audbhav22@iitk.ac.in"]

# Email content
subject = "Application for SURF 2025 Research Internship Opportunity"


def send_email(recipient):
    try:
        # Setup the MIME
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient
        message['Subject'] = subject

        # Attach body text
        message.attach(MIMEText(body, 'plain'))

        # Connect to the server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection

        # Login to the email account
        server.login(sender_email, password)

        # Send the email
        server.sendmail(sender_email, recipient, message.as_string())

        # Quit the server
        server.quit()

        print(f"Email sent to {recipient}")

    except Exception as e:
        print(f"Failed to send email to {recipient}. Error: {e}")
        
    
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Path to the downloaded JSON credentials file
credentials_file = '/Users/udbhavagarwal/Desktop/emailing-439513-a3b6f953c7dd.json'

# Define the scope (permissions required for your application)
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Authorize the client using the credentials and the scope
creds = Credentials.from_service_account_file(credentials_file, scopes=scope)

# Connect to the Google Sheet
client = gspread.authorize(creds)

# Open the Google Sheet by name or IDexi
spreadsheet = client.open("Intern_DB")  # or use open_by_key("<sheet_id>")
sheet = spreadsheet.get_worksheet(1)  # Access the first sheet (use index or name)

expected_headers = ["Professor","Mail Id","Writeup"]
# Read data from the sheet
data = sheet.get_all_records(expected_headers=expected_headers)

df = pd.DataFrame(data)


# Function to add body to the email
def add_body(professor_name, writeup):
    body = f"""
Dear  {professor_name},

{writeup}

Beyond academics, I have honed my leadership and teamwork skills through various roles, such as serving as the Treasurer for the Earth Science Society at IIT Kanpur and the UG Orientation Coordinator for the Earth Science Placement Team. These experiences have equipped me with strong organizational and communication abilities, which I am confident will allow me to contribute effectively to your research group.

I am excited about the opportunity to work under your mentorship, contribute to your research, and deepen my knowledge of geophysical modeling. I have attached my resume for your reference and would appreciate the chance to discuss how my background and skills align with your research goals.

Thank you for your time and consideration. I look forward to the possibility of working with you this summer.

Sincerely,
Shubhangi Jarwal
Department of Earth Sciences
Indian Institute of Technology, Kanpur
"""
    return body

# Function to read sent emails from a file
def read_sent_emails(file_path):
    if os.path.exists(file_path):
        with open(file_path, 'r') as file:
            sent_emails = file.read().splitlines()
    else:
        sent_emails = []
    return sent_emails

# Function to write sent emails to a file
def write_sent_email(file_path, email):
    with open(file_path, 'a') as file:
        file.write(email + '\n')

# File to keep track of sent emails
sent_emails_file = 'sent_emails.txt'
sent_emails = read_sent_emails(sent_emails_file)

# Iterate over the dataframe and send emails
for index, row in df.iterrows():
    recipient = row['Mail Id']
    if recipient not in sent_emails:
        professor_name = row['Professor']
        writeup = row['Writeup']
        body = add_body(professor_name, writeup)
        send_email(recipient)
        write_sent_email(sent_emails_file, recipient)




