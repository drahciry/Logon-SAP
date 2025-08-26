# ------- Imports -------
# Import necessary libraries for the class's functionality.
import win32com.client # Used for COM automation with SAP GUI.
import logging         # Used for creating professional, detailed logs.
import time            # Used for implementing wait logic.

# =============================================================================
# CLASS DEFINITION
# =============================================================================

class SapAutomator:
    """
    A class to encapsulate all SAP GUI automation tasks.
    Each instance of this class will handle one specific logon and task execution.
    """
    def __init__(self, system_name: str, client: int, user: str, password: str):
        """
        Constructor for the SapAutomator class.
        Initializes an instance with all necessary connection details.
        
        Args:
            system_name (str): The description of the SAP system in SAP Logon.
            client (str): The client number for the SAP system.
            user (str): The username for logon.
            password (str): The password for the user.
        """
        # --- Store connection details in instance attributes ---
        self.system_name = system_name
        self.client = client
        self.user = user
        self.password = password
        
        # --- Initialize SAP COM objects to None ---
        # These will be populated when the connect() method is called.
        self.conn = None
        self.session = None
        self.application = None

    def connect(self) -> bool:
        """
        Establishes a connection to the SAP system using the instance's details.
        This method contains the robust wait logic to handle connection delays.
        
        Returns:
            bool: True if the connection and logon are successful, False otherwise.
        """
        # First, check if a password was provided. If not, we cannot proceed.
        if not self.password:
            logging.error(f"No password provided for user {self.user} on system {self.system_name} {self.client}.")
            return False
            
        try:
            # Get the main SAP GUI application object.
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            self.application = sap_gui_auto.GetScriptingEngine
            
            # Check if a connection to the target system already exists.
            print(f"Checking for existing connection to '{self.system_name} {self.client}'...")
            for conn_candidate in self.application.Children:
                if conn_candidate.Description == self.system_name:
                    self.conn = conn_candidate
                    break
            
            # If a connection exists, use it. Otherwise, open a new one.
            if self.conn:
                print(f"Connection to '{self.system_name}' already exists. Reusing it.")
            else:
                print(f"No existing connection found. Opening a new one...")
                self.conn = self.application.OpenConnection(self.system_name, True)
                
                # --- Active Wait Loop ---
                # After requesting a connection, wait actively for the session window to appear.
                # This is more robust than a fixed time.sleep().
                start_time = time.time()
                timeout = 15  # Maximum wait time in seconds.
                while len(self.conn.Children) == 0:
                    time.sleep(1) # Wait for 1 second before checking again.
                    if time.time() - start_time > timeout:
                        print("Timeout exceeded. No session appeared for the connection.")
                        return False
                print("Session window has appeared.")

            # Get the first session object from the connection.
            self.session = self.conn.Children(0)
            
            # Verify that we are on the main logon screen (wnd[0]).
            if not self.session.ActiveWindow.Id.endswith("wnd[0]"):
                logging.error("Active window is not the main logon screen. It might be a popup.")
                return False
                
            logging.info(f"Successfully connected to SAP system: {self.system_name} {self.client}")
            
            # --- Fill Logon Credentials ---
            print("Filling credentials...")
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = self.client
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = self.user
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = self.password
            
            # Send the 'Enter' key to submit the logon information.
            print("Sending logon command...")
            self.session.findById("wnd[0]").sendVKey(0)
            
            time.sleep(2.5) # Short pause to allow logon to be processed.
            
            # --- Handle Post-Logon Popups ---
            try:
                # A popup window is typically identified as wnd[1].
                if self.session.ActiveWindow.Id.endswith("wnd[1]"):
                    print("A popup window was detected after logon. Closing it.")
                    self.session.findById("wnd[1]").close()
            except Exception:
                # If no popup is found, an error will occur, which we can safely ignore.
                pass

            logging.info(f"Login successful for user {self.user} on system {self.system_name} {self.client}.")
            return True # Return True to indicate success.
            
        except Exception as e:
            # Log any unexpected errors that occur during the connection process.
            logging.error(f"An unexpected error occurred during connection: {e}")
            return False # Return False to indicate failure.

    def disconnect(self):
        """
        Disconnects from the SAP session cleanly using the /nex transaction code.
        """
        # Only attempt to disconnect if a session object exists.
        if self.session:
            print("Initiating logoff with /nex...")
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/nex"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1) # Short pause to allow logoff to complete.
            # Reset instance attributes after disconnection.
            self.session = None
            self.conn = None
            print("Logoff command sent.")

    def perform_sample_task(self):
        """
        A placeholder method for the actual business logic automation.
        All transaction navigation and data extraction/entry would go here.
        """
        # Check if we are connected before performing any task.
        if not self.session:
            print("Cannot perform task. Not connected to SAP.")
            return
            
        print("Performing automated tasks...")
        time.sleep(2.5) # This simulates doing some work inside SAP.
        print("Tasks finished.")