# ------- Imports -------
# Import necessary libraries for the script's functionality.
import pandas as pd # Used for reading and handling data from the Excel file.
import logging      # Used for creating professional, detailed logs.
import os           # Used for path manipulation (finding script directory, joining paths).

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def read_all_sheets(file_path) -> dict:
    """
    Reads all sheets from a given Excel file into a dictionary of pandas DataFrames.
    
    Args:
        file_path (str): The full path to the Excel file.
    
    Returns:
        dict: A dictionary where keys are sheet names and values are the corresponding DataFrames.
              Returns None if the file cannot be read.
    """
    try:
        xls = pd.ExcelFile(file_path)
        worksheets_dict = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
        logging.info(f"All sheets from '{os.path.basename(file_path)}' read successfully.")
        return worksheets_dict
    except Exception as e:
        logging.error(f"Error reading Excel file at '{file_path}': {e}")
        return None