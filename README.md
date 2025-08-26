# Logon-SAP: SAP GUI Automation Framework

![Python](https://img.shields.io/badge/Python-3.10+-blue?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-2.3.1-purple?style=for-the-badge&logo=pandas&logoColor=white)
![PyWin32](https://img.shields.io/badge/PyWin32-311-red?style=for-the-badge&logo=pywin32&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

A robust and scalable framework for automating SAP GUI logon and tasks using Python. This project reads connection data from an Excel file, manages credentials securely, and uses an object-oriented design for clean and maintainable code.

## Features

- **Data-Driven Automation**: Reads SAP systems, clients, and users from an Excel spreadsheet, processing all sheets within the file.
- **Secure Credential Management**: Passwords are kept separate from the code in a `credentials.json` file, which should be excluded from version control via `.gitignore`.
- **Object-Oriented Design**: Encapsulates all SAP interaction logic within a reusable `SapAutomator` class.
- **Robust Connection Handling**: Includes active wait logic to handle system delays and ensures a stable connection before proceeding.
- **Professional Logging**: Generates a detailed log file (`sap_automation.log`) for each run, capturing all operations, successes, and errors.
- **Modular & Scalable Structure**: The project is organized with a clear separation of concerns (configuration, functions, models), making it easy to extend and maintain.

## Project Structure

The project follows a professional layout to ensure clarity and scalability:

```
LOGON-SAP/
├── config/              # Holds static configuration files.
│   ├── credentials.json
│   └── credentials_example.json
├── data/                # Contains input data for the script.
│   ├── logon_sap.xlsx
│   └── logon_sap_example.xlsx
├── logs/                # Stores output log files.
│   ├── sap_automation.log
│   └── sap_automation_example.log
├── src/                 # Contains all Python source code.
│   ├── config/          # Python module for configuration logic.
│   │   ├── __init__.py
│   │   └── configs.py
│   ├── functions/       # Python module for helper functions.
│   │   ├── __init__.py
│   │   └── functions.py
│   ├── models/          # Python module for core classes.
│   │   ├── __init__.py
│   │   └── SapAutomator.py
│   └── main.py          # Main entry point of the application.
├── .gitignore
├── LICENSE
├── README.md
└── requirements.txt     # Project dependencies.
```

## Setup and Installation

Follow these steps to set up and run the project locally.

### 1. Prerequisites

- **SAP GUI for Windows**: You must have the SAP GUI client installed on your machine.
- **Enable SAP GUI Scripting**: Scripting must be enabled on both the local client and the server.
    - **Client-Side**: In SAP Logon Options -> Accessibility & Scripting -> Scripting -> Check "Enable scripting". Uncheck "Notify when a script attaches..." and "Notify when a script opens a connection...".
    - **Server-Side**: The system parameter `sapgui/user_scripting` must be set to `TRUE`. Contact your Basis team to confirm this.

### 2. Clone the Repository

Clone this project to your local machine:
```bash
git clone https://github.com/drahciry/Logon-SAP.git
cd Logon-SAP
```

### 3. Create a Virtual Environment

It is highly recommended to use a virtual environment to manage project dependencies.
```bash
# Create the virtual environment
python -m venv venv

# Activate it
# On Windows
.\\venv\\Scripts\\activate
# On macOS/Linux
source venv/bin/activate
```

### 4. Install Dependencies

Install all required Python packages from the `requirements.txt` file.
```bash
pip install -r requirements.txt
```

### 5. Configure Credentials

Your passwords are not stored in the repository. You need to create a local credentials file.
1.  Navigate to the `config/` directory.
2.  Make a copy of `credentials_example.json` and rename it to `credentials.json`.
3.  Edit `credentials.json` and replace the placeholder values with your actual passwords. The structure matches the example file.

### 6. Prepare the Input File

1.  Navigate to the `data/` directory.
2.  Make a copy of `logon_sap_example.xlsx` and rename it to `logon_sap.xlsx`.
3.  Open `logon_sap.xlsx` and fill the sheets with the SAP Systems, Clients, and Users you want to automate. The script will process all sheets in the file.

## Usage

Once the setup is complete, you can run the automation script from the project's root directory:
```bash
python src/main.py
```
The script will print real-time logs to the console and also write a detailed log file to the `logs/` directory.

## How It Works

- **`main.py`**: The main entry point. It orchestrates the entire process by loading credentials and the Excel data, then looping through each entry.
- **`SapAutomator.py`**: An object-oriented class that contains all the logic for interacting with the SAP GUI via the `win32com` library. Each instance handles a single logon/logoff cycle.
- **`configs.py`**: A centralized module for defining file paths and configuring the logging system.
- **`functions.py`**: Contains helper functions, such as `read_all_sheets`, to keep the main script clean.

## Logging

After each run, a detailed log file named `sap_automation.log` is created in the `logs/` directory. This file contains timestamps and information about every step of the process, including successful connections, completed tasks, and any errors that occurred.

## License

This project is licensed under the MIT License - see the `LICENSE` file for details.

## Author

- **LinkedIn** - [Richard Gonçalves](https://linkedin.com/in/drahciry/)
- **GitHub** - [@drahciry](https://github.com/drahciry)
        
