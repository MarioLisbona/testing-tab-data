from lib.utils import setup_google_sheets, process_rfi_cells, write_grouped_rfi_to_sheet

# Example usage
def main():
    client = setup_google_sheets()
    testing_tab = client.open("els-testing-outputs").worksheet("Testing")  # Access the "Testing" tab
    rfi_spreadsheet_tab = client.open("els-testing-outputs").worksheet("RFI Spreadsheet")  # Access the "Testing" tab

    grouped_rfi_entries = process_rfi_cells(testing_tab)  # Call the new function for row 5

    if grouped_rfi_entries:  # Check if there are any entries before writing
        write_grouped_rfi_to_sheet(grouped_rfi_entries, rfi_spreadsheet_tab)


if __name__ == "__main__":
    main()