import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Credentials
GOOGLE_CREDENTIALS = r"C:\Users\Rohit\Downloads\google_credentials.json"

# Scopes
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file"
]

# Option 1: Using URL (Replace with your actual URL)
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/12HI9kh01Umt4tRRqLyklrnomXv2TVGRIL4PnQCMIX88/edit"
WORKSHEET_NAME = "Sheet1"

# Authenticate
creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS, scopes=SCOPES)
client = gspread.authorize(creds)

# Open spreadsheet by URL
spreadsheet = client.open_by_url(SPREADSHEET_URL)

# Or open by name (if in your Drive)
# spreadsheet = client.open("U095 Extracted Data")

# Access worksheet
worksheet = spreadsheet.worksheet(WORKSHEET_NAME)

# Read CSV
df = pd.read_csv("your_data.csv")
data = df.values.tolist()

# Clear existing data
worksheet.clear()

# Update with headers + data
worksheet.update([df.columns.values.tolist()] + data)

print(f"✅ Data uploaded successfully to: {SPREADSHEET_URL}")