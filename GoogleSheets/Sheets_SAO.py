from googleapiclient.errors import HttpError
from googleapiclient.discovery import build


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

#Credentials
CREDENTIALS = "malificent_desk.json"
# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1jIB7OMMWCn63oZlRZyOzysb86Hfr_sekmkfpHKffJ3k"
SAMPLE_RANGE_NAME = "OverallBusiness!A2:A3"

def readSpreadSheet():
  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    values = result.get("values", [])

    if not values:
      print("No data found.")
      return

    print("Name, Major:")
    for row in values:
      # Print columns A and E, which correspond to indices 0 and 4.
      print(f"{row[0]}")
  except HttpError as err:
    print(err)

