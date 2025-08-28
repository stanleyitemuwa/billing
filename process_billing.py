import os
import json
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

def process_google_sheets_data():
    """
    This script replicates the Apps Script logic to update a main Google Sheet
    with data from a newly uploaded sheet. It's designed to be run in GitHub Actions.
    """
    print("🚀 Starting Google Sheet update process...")

    # --- 1. Authenticate and Get Sheet IDs ---
    try:
        print("🔑 Authenticating with Google Cloud...")
        gcp_sa_credentials = json.loads(os.environ["GCP_SA_KEY"])
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(gcp_sa_credentials, scopes=scopes)
        gc = gspread.authorize(creds)

        main_sheet_id = os.environ["MAIN_SHEET_ID"]
        new_data_sheet_id = os.environ["NEW_DATA_SHEET_ID"]
        print("✅ Authentication successful.")
    except Exception as e:
        print(f"❌ Authentication or environment variable error: {e}")
        return

    # --- 2. Load Data from Both Sheets ---
    try:
        print("🔄 Loading data from Google Sheets...")
        # Open main sheet (File A)
        main_sh = gc.open_by_key(main_sheet_id)
        sheet_a = main_sh.worksheet("All_accounts")
        data_a = sheet_a.get_all_values()

        # Open new data sheet (File B)
        new_data_sh = gc.open_by_key(new_data_sheet_id)
        sheet_b = new_data_sh.sheet1
        data_b = sheet_b.get_all_values()
        print("✅ Data loaded successfully.")
    except Exception as e:
        print(f"❌ An error occurred while reading the Google Sheets: {e}")
        return

    # --- 3. Process Data (Replicating Apps Script Logic) ---
    print("⚙️  Processing data: updating existing and appending new rows...")
    
    # Create a lookup dictionary from the main data (File A)
    # Key: Account Number (column 5, index 4), Value: Row Index
    lookup_a = {row[4]: i for i, row in enumerate(data_a[1:], 1)}

    new_rows = []
    
    # Iterate through the new data (File B), skipping the header
    for i, row_b in enumerate(data_b[1:], 1):
        # Account Number is in column 6 (index 5) of the uploaded file
        account_b = row_b[5]
        if not account_b:
            continue

        # Check if the account exists in our main data lookup
        if account_b in lookup_a:
            row_index_a = lookup_a[account_b]
            # Get the data to update (columns 2-5, indices 1-4)
            update_values = row_b[1:5]
            # Replace the first 4 values in the existing row
            data_a[row_index_a][:4] = update_values
        else:
            # If account is new, prepare a new row
            # Get columns 2-8 (indices 1-7) and add the default value 400
            new_row_data = row_b[1:8] + [400]
            new_rows.append(new_row_data)
    
    print(f"Found {len(data_a) - 1 - len(lookup_a)} rows to update and {len(new_rows)} new rows to append.")

    # --- 4. Write Updated Data Back to Google Sheet ---
    try:
        print("✍️ Writing updated data back to the main sheet...")
        # Update the existing rows
        sheet_a.update('A1', data_a, value_input_option='USER_ENTERED')

        # Append the new rows if any exist
        if new_rows:
            sheet_a.append_rows(new_rows, value_input_option='USER_ENTERED')
        
        print("✅ Data update and append complete.")

        # --- 5. Final Cleanup and Formula Insertion ---
        print("🧹 Clearing column L and setting formula...")
        last_row = len(sheet_a.get_all_values())
        if last_row > 1:
            # Clear content in column 12 (L) from row 2 to the end
            range_to_clear = f"L2:L{last_row}"
            sheet_a.batch_clear([range_to_clear])

        # Set the formula in cell L2
        formula = '=ARRAYFORMULA(if(A2:A="","",xlookup(XLOOKUP(B2:B,\'DTR Details\'!L13:L68,\'DTR Details\'!M13:M68),\'DTR Details\'!L5:L9,\'DTR Details\'!M5:M9)))'
        sheet_a.update('L2', formula, value_input_option='USER_ENTERED')
        
        print("✅ Formula set successfully.")

    except Exception as e:
        print(f"❌ An error occurred while writing back to the Google Sheet: {e}")
        return

    # --- 6. Clean up the temporary file ---
    try:
        print(f"🗑️ Deleting temporary sheet (ID: {new_data_sheet_id})...")
        gc.delete_spreadsheet(new_data_sheet_id)
        print("✅ Temporary file deleted.")
    except Exception as e:
        print(f"⚠️ Could not delete temporary file. Manual cleanup may be required. Error: {e}")

    print("\n🎉 All done! Process complete.")


if __name__ == "__main__":
    process_google_sheets_data()
