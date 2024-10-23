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
# sheet1 = spreadsheet.get_worksheet(0)  # Access the first sheet by index
sheet2 = spreadsheet.get_worksheet(1)  # Access the second sheet by index

expected_headers = ["Professor","Mail Id"]
# Read data from the sheet
data = sheet2.get_all_records(expected_headers=expected_headers)

df = pd.DataFrame(data)
print(df)