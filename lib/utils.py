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

# Function to check for "RFI" in a row and update column AJ
def check_rfi_in_row(sheet, row_number):
    row_data = sheet.row_values(row_number)  # Get the values of the specified row
    headers = sheet.row_values(3)  # Get the column headers from row 3
    rfi_found = any("RFI" in cell for cell in row_data if cell)  # Check for "RFI" in non-empty cells

    # Create a list of dictionaries for columns containing "RFI"
    rfi_columns = []
    for index, cell in enumerate(row_data):
        if "RFI" in cell:
            column_reference = index_to_excel_column(index + 1)  # Convert index to Excel-style column reference
            rfi_columns.append({headers[index]: {'value': cell, 'reference': column_reference}})  # Include the column heading and reference

    # Update cell AJ[row_number] based on the presence of "RFI"
    if rfi_found:
        sheet.update(range_name=f'AJ{row_number}', values=[['In Progress - awaiting RFI']])
    else:
        sheet.update(range_name=f'AJ{row_number}', values=[['Yes - no issues']])
    
    return rfi_columns  # Return the list of dictionaries