import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Set up the Google Sheets API
def setup_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('els-testing-sheet-eb7495e198ae.json', scope)
    client = gspread.authorize(creds)
    return client

# Example usage
def main():
    client = setup_google_sheets()
    sheet = client.open("Testing output-MYOM-IHEAB-Testing").worksheet("Testing")  # Access the "Testing" tab

    # Get values from cells E3 and E4
    cell_e3 = sheet.acell('E3').value
    cell_e4 = sheet.acell('E4').value
    cell_aj4 = sheet.acell('AJ4').value
    cell_aj5 = sheet.acell('AJ5').value

    # Log the values
    print("Value in E3:", cell_e3)
    print("Value in E4:", cell_e4)
    print("Value in AJ4:", cell_aj4)
    print("Value in AJ5:", cell_aj5)

    # Update the value in E4 to "test"
    sheet.update(range_name='E4', values=[['test']])  # Use named arguments

    # Confirm the update
    updated_cell_e4 = sheet.acell('E4').value
    print("Updated value in E4:", updated_cell_e4)

if __name__ == "__main__":
    main()
