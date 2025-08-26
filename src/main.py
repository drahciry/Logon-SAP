# =============================================================================
# SAP AUTOMATION SCRIPT
# =============================================================================
#
# Description:
# This script automates the process of logging into multiple SAP systems
# based on data provided in an Excel file. It uses an object-oriented
# approach for clarity and maintainability. Credentials are managed
# securely via a separate `credentials.json` file.
#
# Author: Richard Gonçalves
# Date: 2025-08-26
#
# =============================================================================

# ------- Imports -------
# Import necessary libraries for the script's functionality.
import logging # Used for creating professional, detailed logs.
import json    # Used for reading the secure credentials file.

# Import custom modules.
from models.SapAutomator import SapAutomator
from functions.functions import read_all_sheets
from config.configs import FULL_PATH_TO_EXCEL, FULL_PATH_TO_JSON

# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """
    Main function to orchestrate the entire SAP automation process.
    It loads credentials, reads the Excel file, and iterates through
    each entry to perform the logon and tasks.
    """
    # --- Load Secure Credentials ---
    try:
        # Open and load the JSON file containing the passwords.
        with open(FULL_PATH_TO_JSON) as f:
            credentials = json.load(f)
        logging.info("Secure credentials file 'credentials.json' loaded successfully.")
    except FileNotFoundError:
        logging.critical("credentials.json not found. Please create it. Exiting.")
        exit(1)
    except json.JSONDecodeError:
        logging.critical("Error decoding credentials.json. Please check its format. Exiting.")
        exit(1)

    # ------- Read SAP Logon Data -------
    worksheets = read_all_sheets(FULL_PATH_TO_EXCEL)
    if not worksheets:
        logging.critical('Failed to read Excel file. Exiting script.')
        exit(1)
        
    # ------- Process Each Sheet and Row -------
    for sheet_name, df in worksheets.items():
        if df.empty:
            logging.warning(f"Sheet '{sheet_name}' is empty. Skipping...")
            continue
        
        logging.info(f"Processing sheet: {sheet_name}")
        
        # Iterate over each row in the DataFrame.
        for row in df.itertuples(index=False):
            # Unpack the row data. The Excel file should have these 3 columns.
            sap_system, sap_client, sap_user = row
            
            # --- Look up the Password ---
            # Safely get the password from the loaded credentials dictionary.
            # .get(system, {}) returns an empty dict if the system is not found.
            # .get(user, {}) returns an empty dict if the user is not found for that system.
            # .get(client) returns None if the client is not found for that user.
            password = credentials.get(sap_system, {}).get(sap_user, {}).get(str(sap_client))

            if not password:
                logging.warning(f"Password not found in credentials.json for user '{sap_user}' on system '{sap_system}'. Skipping.")
                continue

            # --- Create and Use the Automator Object ---
            # Create an instance of our SapAutomator class for this specific entry.
            automator = SapAutomator(
                system_name=sap_system, 
                client=str(sap_client), # Ensure client is a string
                user=sap_user,
                password=password
            )
            
            # Execute the automation flow: connect, perform task, disconnect.
            if automator.connect():
                automator.perform_sample_task()
                automator.disconnect()
            else:
                logging.error(f"Failed to complete the process for user {sap_user} on system {sap_system} {sap_client}.")
    
    logging.info("Script finished processing all entries.")

# ------- Script Entry Point -------
# The __name__ == "__main__" guard ensures that the main() function is called
# only when the script is executed directly (not when imported as a module).
if __name__ == "__main__":
    main()