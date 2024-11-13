import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json  # Import json module to handle JSON file creation

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

# Function to process RFI cells in all rows with data
def process_rfi_cells(sheet):
    all_values = sheet.get_all_values()  # Get all values in the sheet
    headers = sheet.row_values(3)  # Get the column headers from row 3
    rfi_entries = []  # Initialize list to store RFI entries

    for row_number in range(4, len(all_values) + 1):  # Start from row 4 to the last row
        row_data = sheet.row_values(row_number)  # Get the values of the specified row

        # Returns a boolean if any non-empty cells are found with the "RFI" substring - Exclude AJ (column 35)
        rfi_found = any("RFI" in cell for index, cell in enumerate(row_data) if index != 35 and cell)  

        # Get the Implementation Identifier from header index C (index 2)
        implementation_identifier = row_data[2] if len(row_data) > 2 else None  # Ensure the row has enough data

        # Create a list of dictionaries for columns containing "RFI", excluding column AJ
        for index, cell in enumerate(row_data):
            if index != 35 and "RFI" in cell:  # Exclude column AJ (index 35)
                column_reference = index_to_excel_column(index + 1)  # Convert index to Excel-style column reference
                rfi_entries.append({
                    "column_header": headers[index],
                    "reference": column_reference,
                    "cell": f'{index_to_excel_column(index + 1)}{row_number}',  # Include the cell location
                    "value": cell,
                    "implementation_identifier": implementation_identifier  # Include the Implementation Identifier
                })  # Include the column heading and reference

        # Update cell AJ[row_number] based on the presence of "RFI"
        if rfi_found:
            sheet.update(range_name=f'AJ{row_number}', values=[['In Progress - awaiting RFI']])
        else:
            sheet.update(range_name=f'AJ{row_number}', values=[['Yes - no issues']])
    
    # Group RFI entries by their value
    grouped_rfi_entries = {}
    for entry in rfi_entries:
        value = entry["value"]
        if value not in grouped_rfi_entries:
            grouped_rfi_entries[value] = {
                "RFI_value": value,
                "projects_effected": []
            }
        grouped_rfi_entries[value]["projects_effected"].append({
            "reference": entry["reference"],
            "cell": entry["cell"],
            "implementation_identifier": entry["implementation_identifier"]
        })

    # Create rfi.json in the root directory
    with open('rfi.json', 'w', encoding='utf-8') as json_file:
        json.dump({"RFI_Data": list(grouped_rfi_entries.values())}, json_file, indent=4)  # Write the grouped RFI entries to the JSON file
    
    return grouped_rfi_entries  # Return the list of dictionaries

# Function to write grouped RFI entries to a Google Sheet
def write_grouped_rfi_to_sheet(grouped_rfi_entries, sheet):
    # Define the starting row after the header (row 6)
    start_row = 7  # Row 7 is the first empty row after row 6

    # Get all values from the sheet starting from row 7
    existing_values = sheet.get_all_values()  # Fetch all values at once

    # Prepare lists to hold the values for batch update
    rfi_values = []
    implementation_identifiers = []

    for entry in grouped_rfi_entries.values():  # Iterate over the values of the dictionary
        rfi_value = entry["RFI_value"]
        projects_effected = entry["projects_effected"]

        # Only proceed if the number of projects_effected is greater than 5
        # and the RFI_value is not "In Progress - awaiting RFI"
        if len(projects_effected) >= 5 and rfi_value != "In Progress - awaiting RFI":
            # Find the first empty row after row 6
            empty_row = start_row
            while empty_row < len(existing_values) and existing_values[empty_row - 1]:  # Check if the row is not empty
                empty_row += 1

            # Append values to the lists for batch update
            rfi_values.append([rfi_value])
            identifiers = [project["implementation_identifier"] for project in projects_effected]
            implementation_identifiers.append([", ".join(identifiers)])  # Join identifiers with a comma

    # Perform batch update for RFI values
    if rfi_values:
        sheet.update(range_name=f'A{start_row}:{chr(65 + len(rfi_values[0]) - 1)}{start_row + len(rfi_values) - 1}', values=rfi_values)

    # Perform batch update for implementation identifiers
    if implementation_identifiers:
        sheet.update(range_name=f'B{start_row}:{chr(66 + len(implementation_identifiers[0]) - 1)}{start_row + len(implementation_identifiers) - 1}', values=implementation_identifiers)


