import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Set up the Google Sheets API
def setup_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('els-testing-sheet-eb7495e198ae.json', scope)
    client = gspread.authorize(creds)
    return client

# Function to log the number of rows with data and the first empty row
def log_row_data(sheet):
    all_values = sheet.get_all_values()  # Get all values in the sheet
    row_count = len(all_values)  # Count the total number of rows
    print("Total rows with data:", row_count)

    # Find the first empty row
    first_empty_row = None
    for i, row in enumerate(all_values):
        if all(cell == '' for cell in row):  # Check if the entire row is empty
            first_empty_row = i + 1  # Row numbers are 1-indexed
            break

    if first_empty_row:
        print("First empty row is:", first_empty_row)
    else:
        print("No empty rows found.")

# Helper function to convert column index to Excel-style notation
def index_to_excel_column(index):
    column = ""
    while index > 0:
        index -= 1
        column = chr(index % 26 + 65) + column  # Convert to ASCII and prepend
        index //= 26
    return column

# Function to log the last column with data
def log_last_column_with_data(sheet):
    all_values = sheet.get_all_values()  # Get all values in the sheet
    last_column_index = 0  # Initialize the last column index

    # Iterate through each row to find the last column with data
    for row in all_values:
        for index, cell in enumerate(row):
            if cell:  # If the cell is not empty
                last_column_index = max(last_column_index, index + 1)  # Update the last column index (1-indexed)

    if last_column_index > 0:
        excel_column = index_to_excel_column(last_column_index)  # Convert to Excel-style column
        print("Last column with data is:", excel_column)
    else:
        print("No data found in the sheet.")

# Example usage
def main():
    client = setup_google_sheets()
    sheet = client.open("Testing output-MYOM-IHEAB-Testing").worksheet("Testing")  # Access the "Testing" tab

    # Log the number of rows with data and the first empty row
    log_row_data(sheet)

    # Log the last column with data
    log_last_column_with_data(sheet)

    # Get values from cells E3 and E4
    cell_e3 = sheet.acell('E3').value
    cell_e4 = sheet.acell('E4').value

    # Log the values
    print("Value in E3:", cell_e3)
    print("Value in E4:", cell_e4)

    # Update the value in E4 to "test"
    sheet.update(range_name='E4', values=[['test']])  # Use named arguments

    # Confirm the update
    updated_cell_e4 = sheet.acell('E4').value
    print("Updated value in E4:", updated_cell_e4)

if __name__ == "__main__":
    main()
