import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Set up the Google Sheets API
def setup_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('els-testing-sheet-eb7495e198ae.json', scope)
    client = gspread.authorize(creds)
    return client

# Import a row from Google Sheets
def import_row(sheet, row_number):
    return sheet.row_values(row_number)

# Update a row in Google Sheets
def update_row(sheet, row_number, updated_data):
    sheet.update(f'A{row_number}', [updated_data])

# Example usage
def main():
    client = setup_google_sheets()
    sheet = client.open("Testing output-MYOM-IHEAB-Testing").worksheet("Testing")  # Access the "Testing" tab

    row_number = 2  # Change this to the row you want to import
    row_data = import_row(sheet, row_number)

    # Make some changes to the row data
    row_data[4] = "Testing some updated values"  # Example change to the first column

    # Update the row back to the sheet
    update_row(sheet, row_number, row_data)

if __name__ == "__main__":
    main()
