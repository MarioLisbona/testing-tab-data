from lib.utils import setup_google_sheets, check_rfi_in_row




# Example usage
def main():
    client = setup_google_sheets()
    sheet = client.open("Testing output-MYOM-IHEAB-Testing").worksheet("Testing")  # Access the "Testing" tab

    # Check RFI in a specific row (for example, row 5)
    # print(check_rfi_in_row(sheet, 4))  # Call the new function for row 5
    print(check_rfi_in_row(sheet, 25))  # Call the new function for row 5

if __name__ == "__main__":
    main()
