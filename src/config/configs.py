# ------- Imports -------
# Import necessary libraries for the script's functionality.
import logging          # Used for creating professional, detailed logs.
import os               # Used for path manipulation (finding script directory, joining paths).

# ------- Configuration -------
# This section sets up global configurations like logging and paths.

# Determine the absolute path of the script and its parent directory.
# This makes the script portable and independent of where it is run from.
config_dir = os.path.dirname(os.path.abspath(__file__))
source_dir = os.path.dirname(config_dir)
project_root = os.path.dirname(source_dir)

# --- Setup professional logging ---
# Ensure the 'logs' directory exists. If not, create it.
log_dir = os.path.join(project_root, "logs")
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

# Configure the logging module.
logging.basicConfig(
    level=logging.INFO,  # Set the minimum level of messages to log (e.g., INFO, WARNING, ERROR).
    format='%(asctime)s - %(levelname)s - %(message)s',  # Define the format for log messages.
    filename=os.path.join(log_dir, "sap_automation.log"),  # Specify the log file name and path.
    filemode='w'  # 'w' overwrites the log file on each run. Use 'a' to append to the log.
)

# Add a handler to also print logs to the console in real-time.
# This is useful for monitoring the script's execution.
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)  # Set the level for the console output.
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)
logging.getLogger().addHandler(console_handler)

# ------- Constants -------
# Define constants used throughout the script.
FULL_PATH_TO_EXCEL = os.path.join(project_root, "data", "logon_sap.xlsx")
FULL_PATH_TO_JSON = os.path.join(project_root, "config", "credentials.json")