import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from email.mime.application import MIMEApplication  # Add this import


# Load environment variables
load_dotenv()

smtp_server = os.getenv('SMTP_SERVER')
smtp_port = int(os.getenv('SMTP_PORT', '25'))
sender_email = os.getenv('SENDER_EMAIL')
password = os.getenv('EMAIL_PASSWORD')

# recipients = ["shubhangij22@iitk.ac.in", "audbhav22@iitk.ac.in"]

# Email content
subject = "Application for SURF 2025 Research Internship Opportunity"


def send_email(recipient, attachment_path=None):
    try:
        # Setup the MIME
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient
        message['Subject'] = subject

        # Attach body text
        message.attach(MIMEText(body, 'plain'))

        # Attach file if provided
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as f:
                attachment = MIMEApplication(f.read(), _subtype='pdf')
                attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                message.attach(attachment)

        # Connect to the server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection

        # Login 
        server.login(sender_email, password)

        # Send 
        server.sendmail(sender_email, recipient, message.as_string())

        # Quit 
        server.quit()

        print(f"Email sent to {recipient}")

    except Exception as e:
        print(f"Failed to send email to {recipient}. Error: {e}")
        


# Google credentials file
credentials_file = '/Users/udbhavagarwal/Desktop/emailing-439513-a3b6f953c7dd.json'

# scope (permissions required for your application)
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Authorize the client 
creds = Credentials.from_service_account_file(credentials_file, scopes=scope)
client = gspread.authorize(creds)

spreadsheet = client.open("Intern_DB")  # or use open_by_key("<sheet_id>")
sheet = spreadsheet.get_worksheet(1)  # Access by worksheet 

expected_headers = ["Professor","Mail Id","Writeup"]

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

# read sent emails from a file
def read_sent_emails(file_path):
    if os.path.exists(file_path):
        with open(file_path, 'r') as file:
            sent_emails = file.read().splitlines()
    else:
        sent_emails = []
    return sent_emails

#  write sent emails to a file
def write_sent_email(file_path, email):
    with open(file_path, 'a') as file:
        file.write(email + '\n')

# keep track of sent emails
sent_emails_file = 'sent_emails.txt'
sent_emails = read_sent_emails(sent_emails_file)

# Iterate over the dataframe and send emails
for index, row in df.iterrows():
    recipient = row['Mail Id']
    if recipient not in sent_emails:
        professor_name = row['Professor']
        writeup = row['Writeup']
        body = add_body(professor_name, writeup)
        send_email(recipient, attachment_path='path/to/your/resume.pdf')  # Add your resume path here
        write_sent_email(sent_emails_file, recipient)




