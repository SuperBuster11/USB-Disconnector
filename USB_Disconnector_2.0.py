import os
import subprocess
import ctypes
import sys
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import threading
import tempfile
import json
import re # Import for regular expressions to extract VID/PID
import traceback # Import for traceback module

# --- Global Exception Handler ---
def global_exception_handler(exc_type, exc_value, exc_traceback):
    """Global exception handler to catch all uncaught exceptions."""
    # Print the error to console
    print("\n" + "="*50)
    print("UNCAUGHT EXCEPTION!")
    print("="*50)
    traceback.print_exception(exc_type, exc_value, exc_traceback)
    print("="*50 + "\n")

    # Show a messagebox to the user
    error_message = f"Wystąpił nieoczekiwany błąd w aplikacji:\n\n" \
                    f"Typ błędu: {exc_type.__name__}\n" \
                    f"Komunikat: {exc_value}\n\n" \
                    f"Szczegóły błędu w konsoli."
    try:
        # Try to use Tkinter messagebox if Tkinter is already initialized
        # This might fail if the crash happens too early before root is created
        messagebox.showerror("Krytyczny Błąd Aplikacji", error_message)
    except tk.TclError:
        # If Tkinter is not ready, just print and exit.
        print(f"Could not show messagebox: Tkinter not ready. Error: {error_message}")
    finally:
        # Ensure the program exits after reporting the error
        sys.exit(1)

sys.excepthook = global_exception_handler
# --- End Global Exception Handler ---


# --- Device Name Database Configuration ---
DEVICE_DATABASE_FILE = "device_names.json"

def load_device_database():
    """
    Loads the device name database from a JSON file.
    The database maps "VID:PID" keys to human-readable names.
    """
    if os.path.exists(DEVICE_DATABASE_FILE):
        try:
            with open(DEVICE_DATABASE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            print(f"Warning: Could not decode {DEVICE_DATABASE_FILE}. Creating a new empty database.")
            return {}
        except Exception as e:
            print(f"Error loading {DEVICE_DATABASE_FILE}: {e}. Creating a new empty database.")
            return {}
    return {} # Return empty dict if file doesn't exist

def save_device_database(db):
    """
    Saves the device name database to a JSON file.
    """
    try:
        with open(DEVICE_DATABASE_FILE, 'w', encoding='utf-8') as f:
            json.dump(db, f, indent=4) # Use indent for human-readable JSON
    except Exception as e:
        print(f"Error saving {DEVICE_DATABASE_FILE}: {e}")

def extract_vid_pid(id_string):
    """
    Extracts Vendor ID (VID) and Product ID (PID) from a string like InstanceId or HardwareId.
    Expected format: VID_XXXX&PID_YYYY
    """
    vid_match = re.search(r'VID_([0-9A-Fa-f]{4})', id_string, re.IGNORECASE)
    pid_match = re.search(r'PID_([0-9A-Fa-f]{4})', id_string, re.IGNORECASE)
    
    vid = vid_match.group(1).upper() if vid_match else None
    pid = pid_match.group(1).upper() if pid_match else None
    
    return vid, pid

# --- End Device Name Database Configuration ---


def get_resource_path(relative_path):
    """
    Get the absolute path to a resource, works for dev and for PyInstaller.
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def run_as_admin(command_args, use_powershell=False):
    """
    Attempts to run a command with administrator privileges.
    If the script is not already running as admin, it will try to re-launch itself with elevated privileges.
    For PowerShell commands, it redirects output to a temporary file to ensure capture.
    """
    if ctypes.windll.shell32.IsUserAnAdmin(): # This was 'IsUserAnAdmin' in provided code, should be 'IsUserAdmin' or 'IsUserAnAdmin'
        temp_file_path = None # Initialize to None for cleanup in finally block

        try:
            encoding_to_use = 'cp1250' # Standard encoding for Polish Windows console output

            # Configure startupinfo to show the console window normally.
            # This constant is part of the Windows API (ShowWindow). SW_NORMAL is 1.
            SW_NORMAL = 1 

            startupinfo = None
            if sys.platform == "win32":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW # Indicate that wShowWindow should be used
                startupinfo.wShowWindow = SW_NORMAL # Show the window normally

            if use_powershell:
                # Create a temporary file to capture PowerShell output
                # Using NamedTemporaryFile with delete=False allows reading after process completion
                with tempfile.NamedTemporaryFile(mode='w+', delete=False, encoding=encoding_to_use, suffix=".txt") as temp_output_file:
                    temp_file_path = temp_output_file.name

                # Construct the PowerShell command to redirect its output to the temporary file.
                # 'Out-File -FilePath' is used for reliable redirection in PowerShell.
                # '-Encoding Default' matches system's default encoding (usually cp1250 on Polish Windows).
                powershell_command_with_redirect = f"{command_args} | Out-File -FilePath '{temp_file_path}' -Encoding Default"
                
                # Execute PowerShell command. capture_output=True still captures stderr from PowerShell itself.
                process = subprocess.run(['powershell', '-command', powershell_command_with_redirect],
                                         capture_output=True, text=True, encoding=encoding_to_use, check=True, shell=False,
                                         startupinfo=startupinfo)
                
                # Read the output from the temporary file.
                # Seek to the beginning of the file before reading.
                with open(temp_file_path, 'r', encoding=encoding_to_use) as f:
                    stdout_content = f.read()
                
                # PowerShell's actual errors (e.g., syntax errors) might still be in process.stderr
                stderr_content = process.stderr
                returncode = process.returncode

                return stdout_content, stderr_content, returncode

            else:
                # For non-PowerShell commands (like pnputil), continue with direct capture.
                # No file redirection needed here.
                process = subprocess.run(command_args, capture_output=True, text=True, encoding=encoding_to_use, check=True, shell=False,
                                         startupinfo=startupinfo)
                return process.stdout, process.stderr, process.returncode

        except subprocess.CalledProcessError as e:
            # In case of CalledProcessError, still try to read from temp file if it exists
            stdout_from_temp = ""
            if temp_file_path and os.path.exists(temp_file_path):
                try:
                    with open(temp_file_path, 'r', encoding=encoding_to_use) as f:
                        stdout_from_temp = f.read()
                except Exception as file_read_e:
                    print(f"Error reading temp file after CalledProcessError: {file_read_e}")
            return stdout_from_temp, e.stderr if e.stderr else "Nieznany błąd.", e.returncode
        except FileNotFoundError:
            # Handle cases where the command executable is not found
            return "", "Program (PowerShell lub pnputil.exe) nie został znaleziony. Upewnij się, że jest w PATH.", -1
        except Exception as e:
            # Catch any other unexpected errors during subprocess execution
            return "", f"Wystąpił nieoczekiwany błąd subprocess: {e}", -1
        finally:
            # Ensure the temporary file is deleted in all cases
            if temp_file_path and os.path.exists(temp_file_path):
                try:
                    os.remove(temp_file_path)
                except Exception as cleanup_e:
                    print(f"Error cleaning up temporary file '{temp_file_path}': {cleanup_e}")
    else:
        # If not running as admin, re-launch the script with admin privileges
        try:
            # ShellExecuteW with "runas" verb attempts to elevate privileges.
            # The last argument (1) means SW_SHOWNORMAL, so the new window will be visible.
            # Ensure __file__ path is correctly quoted to handle spaces.
            script_path = os.path.abspath(sys.argv[0]) # Get path to current script
            quoted_script_path = f'"{script_path}"' # Ensure proper quoting for paths with spaces

            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, quoted_script_path, None, 1)
            sys.exit(0) # Exit the current non-elevated process
        except Exception as e:
            # Inform the user if elevation failed
            messagebox.showerror("Błąd uruchamiania", f"Nie udało się uruchomić programu z uprawnieniami administratora.\n{e}")
            sys.exit(1)


def get_device_status(device_id):
    """
    Determines the current status (enabled/disabled/unknown) of a given device ID
    by querying its properties (ConfigManagerErrorCode, Status, Present) using PowerShell,
    and parsing the output as JSON for robustness.
    """
    # Query multiple properties in one go and convert to JSON for robust parsing
    ps_command = (
        f"Get-PnpDevice -InstanceId '{device_id}' | "
        f"Select-Object ConfigManagerErrorCode, Status, Present | ConvertTo-Json -Compress"
    )
    stdout, stderr, returncode = run_as_admin(ps_command, use_powershell=True)

    if returncode == 0 and stdout:
        try:
            device_props = json.loads(stdout)
            
            # ConvertTo-Json might return a single object or an array of objects.
            # We expect a single object for a single InstanceId query.
            if isinstance(device_props, list) and len(device_props) > 0:
                device_props = device_props[0]
            elif not isinstance(device_props, dict):
                # If it's not a dictionary (or a list containing one), it's an unexpected format.
                print(f"DEBUG: Nieoczekiwany format JSON dla statusu urządzenia '{device_id}': {device_props}")
                return "unknown"
            
            # --- Prioritize 'Present' status for disabled devices ---
            # If a device is explicitly not present, it's considered disabled (e.g., by pnputil /disable-device)
            present_status = device_props.get("Present")
            if present_status is False: # Explicitly checking for boolean False
                return "disabled"

            # --- Fallback to ConfigManagerErrorCode and Status property ---
            config_manager_error_code = str(device_props.get("ConfigManagerErrorCode", "")).strip()
            status_text = str(device_props.get("Status", "")).strip().upper()

            if config_manager_error_code == "22": # Code 22 often means device is disabled
                return "disabled"
            if config_manager_error_code == "0": # Code 0 means device is working correctly
                return "enabled"
            if "OK" in status_text: # General 'OK' status indicates enabled
                return "enabled"
            
            return "unknown" # Default if no specific status matches

        except json.JSONDecodeError as e:
            messagebox.showerror("Błąd Parsowania JSON", f"Nie udało się sparsować odpowiedzi PowerShell dla statusu.\nBłąd: {e}\n\nSurowe wyjście:\n{stdout}")
            print(f"DEBUG: JSON Decoding Error w get_device_status: {e}\nRaw STDOUT: '{stdout}'") # Debug print for console
            return "unknown"
        except Exception as e:
            messagebox.showerror("Błąd Danych Statusu", f"Wystąpił nieoczekiwany błąd podczas parsowania statusu urządzenia '{device_id}': {e}.\n\nSurowe wyjście:\n{stdout}")
            print(f"DEBUG: Nieoczekiwany błąd w parsowaniu statusu: {e}\nRaw STDOUT: '{stdout}'") # Debug print for console
            return "unknown"
    else:
        # Debugging information if command fails or no output
        print(f"DEBUG: PowerShell command output dla statusu '{device_id}':")
        print(f"  Return Code: {returncode}")
        print(f"  STDOUT (raw): '{stdout}'")
        print(f"  STDERR (raw): '{stderr}'")
        # If no stdout and return code is 0, it means no device found with that InstanceId or no properties were returned.
        # This can happen if a disabled device is sometimes not considered "present" even if it's not removed.
        if not stdout and returncode == 0:
            print(f"DEBUG: Brak danych statusu dla {device_id}. Może być usunięte lub bardzo ukryte.")
        return "unknown" # Return unknown if command failed or returned no data


def toggle_device_status(device_id, action):
    """
    Toggles the status of a device (enable/disable) using pnputil.exe.
    Displays success or error messages to the user.
    """
    command = ["pnputil", f"/{action}-device", device_id]
    stdout, stderr, returncode = run_as_admin(command, use_powershell=False)

    if returncode == 0:
        message = f"Pomyślnie {action}owano urządzenie:\n{device_id}"
        if stdout:
            message += f"\n\nSzczegóły:\n{stdout.strip()}"
        messagebox.showinfo("Sukces", message)
    else:
        error_message = f"Wystąpił błąd podczas {action}owania urządzenia:\n{device_id}"
        error_message += f"\n\nKod błędu: {returncode}"
        if stdout:
            error_message += f"\nSTDOUT: {stdout.strip()}"
            # Only append STDOUT if it's not empty, otherwise it adds a blank line.
        if stderr:
            error_message += f"\nSTDERR: {stderr.strip()}"
        messagebox.showerror("Błąd Operacji", error_message)

    # After toggling, update the UI status if the app instance exists
    if UsbDeviceControllerApp.instance:
        # Call update_selected_device_status to refresh UI after toggle
        UsbDeviceControllerApp.instance.update_selected_device_status()


def get_input_devices_by_instance_id_pattern(device_database): # Now accepts device_database
    """
    Retrieves a list of USB input devices (HID, Keyboard, Mouse, PointingDevice, GamePort)
    using PowerShell's Get-PnpDevice and converts the output to JSON for robust parsing.
    It also uses the provided device_database to get custom names.
    """
    ps_command = (
        r"Get-PnpDevice -PresentOnly | " # Only get devices currently connected
        r"Where-Object { ($_.InstanceId -like 'USB\VID_*&PID_*\*') -and " # Filter for USB devices with VID/PID pattern
        r"($_.Class -eq 'HIDClass' -or $_.Class -eq 'Keyboard' -or $_.Class -eq 'Mouse' -or $_.Class -eq 'PointingDevice' -or $_.Class -eq 'GamePort') } | " # Filter by specific device classes
        r"Select-Object DeviceDescription, FriendlyName, InstanceId, HardwareId | ConvertTo-Json -Compress" # Select and format as JSON
    )
    
    stdout, stderr, returncode = run_as_admin(ps_command, use_powershell=True)

    devices_list = []
    seen_ids = set() # To track unique InstanceIds to prevent duplicates
    if returncode == 0 and stdout:
        try:
            json_data = json.loads(stdout)
            # Ensure json_data is a list, even if ConvertTo-Json returns a single object
            if not isinstance(json_data, list):
                json_data = [json_data] # Wrap single object in a list

            for dev_info in json_data:
                instance_id = dev_info.get("InstanceId")
                
                # --- Fix for Duplicate Entries: Skip if this InstanceId has already been processed ---
                if instance_id in seen_ids:
                    print(f"DEBUG: Pomijanie duplikatu urządzenia InstanceId: {instance_id}")
                    continue
                seen_ids.add(instance_id)
                # --- End Fix ---

                hardware_ids = dev_info.get("HardwareId", []) # Get HardwareId, it can be a list or string
                
                # Ensure hardware_ids is always a list for iteration
                if isinstance(hardware_ids, str):
                    hardware_ids = [hardware_ids]
                
                vid, pid = None, None
                
                # Try extracting VID/PID from InstanceId first
                if instance_id:
                    vid, pid = extract_vid_pid(instance_id)

                # If not found, try extracting from HardwareId(s)
                if not (vid and pid) and hardware_ids:
                    for hwid_entry in hardware_ids:
                        vid, pid = extract_vid_pid(hwid_entry)
                        if vid and pid:
                            break # Found VID/PID in one of the HardwareIds

                device_key = f"{vid}:{pid}" if (vid and pid) else None
                custom_name = None
                if device_key:
                    custom_name = device_database.get(device_key) # Lookup in the provided database

                final_display_name = ""
                if custom_name:
                    final_display_name = custom_name
                else:
                    # Fallback to existing logic if no custom name
                    final_display_name = dev_info.get("FriendlyName") or dev_info.get("DeviceDescription") or instance_id

                # Append VID:PID to the display name for easy identification, even if custom name exists.
                if device_key:
                    final_display_name = f"{final_display_name} [{device_key}]"

                # InstanceId is crucial, ensure it's present before adding to list
                if instance_id:
                    devices_list.append({
                        "display_name": final_display_name,
                        "id": instance_id, # This is the full InstanceId used for pnputil
                        "hardware_id": hardware_ids, # Keep original hardware_ids
                        "vid_pid_key": device_key, # Store the VID:PID key for potential future use
                        "full_info": dev_info
                    })
        except json.JSONDecodeError as e:
            messagebox.showerror("Błąd Parsowania JSON", f"Nie udało się sparsować odpowiedzi PowerShell.\nBłąd: {e}\n\nSurowe wyjście:\n{stdout}")
            print(f"DEBUG: JSON Decoding Error: {e}\nRaw STDOUT: '{stdout}'") # Debug print for console
        except KeyError as e:
            messagebox.showerror("Błąd Danych Urządzenia", f"Brak oczekiwanego klucza w danych urządzenia ({e}).\n\nSurowe dane (dev_info): {dev_info}")
            print(f"DEBUG: Missing Key Error: {e}\nRaw dev_info: {dev_info}") # Debug print for console
    else:
        # Debugging information when no devices are found or command fails
        print(f"DEBUG: PowerShell command output:")
        print(f"  Return Code: {returncode}")
        print(f"  STDOUT (raw): '{stdout}'")
        print(f"  STDERR (raw): '{stderr}')")
        if not stdout and returncode == 0:
            messagebox.showerror("Błąd", f"Nie udało się pobrać listy urządzeń wejściowych: Komenda PowerShell zwróciła pusty wynik.\n"
                                          f"Sprawdź, czy są podłączone urządzenia wejściowe USB (klawiatury, myszy, gamepady itp.) spełniające kryteria.")
        else:
            messagebox.showerror("Błąd", f"Nie udało się pobrać listy urządzeń wejściowych:\n{stderr if stderr else 'Nieznany błąd.'}")


    devices_list.sort(key=lambda x: x["display_name"].lower()) # Sort devices alphabetically
    return devices_list


class UsbDeviceControllerApp:
    """
    Main application class for the Tkinter GUI.
    Manages device selection, status display, and enable/disable operations.
    """
    instance = None # Class-level variable to hold the single instance of the app

    def __init__(self, master):
        print("DEBUG: UsbDeviceControllerApp __init__ started.") # DEBUG PRINT
        UsbDeviceControllerApp.instance = self # Assign the current instance to the class variable
        self.master = master
        master.title("USB-Disconnector v2.0")
        master.geometry("700x350")
        master.resizable(False, False) # Prevent window resizing
        
        # --- Make main window always on top ---
        master.attributes('-topmost', True) 

        # --- Add window icon using resource path ---
        try:
            # Get the correct path to the icon file, working for both dev and compiled environments.
            icon_path = get_resource_path('type-a.ico')
            master.iconbitmap(icon_path)
        except tk.TclError:
            # Print a warning if the icon file cannot be found or loaded
            print(f"Warning: Could not load icon '{icon_path}'. Make sure it's in the correct location and bundled with PyInstaller.")
        # --- End window icon ---

        # --- Load device name database ---
        self.device_db = load_device_database()

        # --- Create a main frame to contain all UI elements ---
        self.main_container_frame = ttk.Frame(master, padding=10)
        self.main_container_frame.pack(fill="both", expand=True)

        # Frame for device selection (now child of main_container_frame)
        self.device_selection_frame = ttk.LabelFrame(self.main_container_frame, text="Wybierz urządzenie wejściowe", padding=(10, 10))
        self.device_selection_frame.grid(row=0, column=0, padx=0, pady=0, sticky="nsew", columnspan=2) 

        # Device selection combobox and Edit button within the frame
        tk.Label(self.device_selection_frame, text="Urządzenie:", font=("Arial", 11)).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.selected_device_name = tk.StringVar()
        self.device_combobox = ttk.Combobox(self.device_selection_frame, textvariable=self.selected_device_name, width=50, state="readonly", font=("Arial", 10)) 
        self.device_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew") 
        self.device_combobox.bind("<<ComboboxSelected>>", self.on_device_selected)


        # --- Edit Name button (positioned next to combobox) ---
        self.edit_name_button = tk.Button(self.device_selection_frame, text="Edytuj Nazwę", # Shorter text for button
                                           command=self._open_edit_name_dialog,
                                           bg="#87CEEB", fg="black", width=15, height=1, font=("Arial", 9, "bold"), relief="raised", state="disabled") 
        self.edit_name_button.grid(row=0, column=2, padx=5, pady=5, sticky="e") # Placed in new column 2

        # Status indicator (colored circle) and text label - now span two columns
        self.status_indicator = tk.Label(self.device_selection_frame, text="●", font=("Arial", 18, "bold"), width=2)
        self.status_indicator.grid(row=1, column=0, padx=5, pady=10, sticky="w")
        self.status_label_text = tk.StringVar(value="Status: Nieznany")
        self.status_text_label = tk.Label(self.device_selection_frame, textvariable=self.status_label_text, font=("Arial", 11))
        # Spans the remaining columns (column 1 and 2)
        self.status_text_label.grid(row=1, column=1, padx=5, pady=10, sticky="w", columnspan=2) 

        # --- Frame to hold the main action buttons (Disable, Enable, Refresh) ---
        self.main_action_buttons_frame = ttk.Frame(self.main_container_frame) # Now child of main_container_frame
        # This frame is placed in column 0 of the main_container_frame, spanning rows 1-3.
        # sticky="ns" ensures vertical centering within its grid cell.
        self.main_action_buttons_frame.grid(row=1, column=0, rowspan=3, padx=0, pady=0, sticky="ns", columnspan=2) # Spans both columns of the main_container_frame


        # Action buttons are now packed inside the new frame.
        # .pack() without 'side' defaults to TOP, stacking them vertically.
        # Removed 'sticky="ew"' from buttons to allow them to be centered within the frame.
        self.disable_button = tk.Button(self.main_action_buttons_frame, text="Wyłącz wybrane urządzenie",
                                         command=self.disable_selected_device,
                                         bg="#FF6666", fg="white", width=25, height=2, font=("Arial", 10, "bold"), relief="raised", state="disabled")
        self.disable_button.pack(pady=5) 

        self.enable_button = tk.Button(self.main_action_buttons_frame, text="Włącz wybrane urządzenie",
                                        command=self.enable_selected_device,
                                        bg="#66FF66", fg="black", width=25, height=2, font=("Arial", 10, "bold"), relief="raised", state="disabled")
        self.enable_button.pack(pady=5) 

        self.refresh_button = tk.Button(self.main_action_buttons_frame, text="Odśwież listę urządzeń",
                                         command=self.populate_devices,
                                         bg="#ADD8E6", fg="black", width=25, height=2, font=("Arial", 10, "bold"), relief="raised")
        self.refresh_button.pack(pady=5) 

        # Admin privilege reminder - now child of main_container_frame
        tk.Label(self.main_container_frame, text="Pamiętaj: program wymaga uprawnień administratora.", font=("Arial", 8, "italic"), fg="gray") \
            .grid(row=4, column=0, pady=(10, 5), columnspan=2) 

        # Configure grid weights for the main_container_frame
        self.main_container_frame.grid_rowconfigure(0, weight=1) # Allows device_selection_frame to expand
        self.main_container_frame.grid_columnconfigure(0, weight=1) # Allows the action_buttons_frame column to expand (and center its content)
        self.main_container_frame.grid_columnconfigure(1, weight=1) # Allows the right side of the main_container_frame to expand

        # Configure internal grid weights for device_selection_frame
        self.device_selection_frame.grid_columnconfigure(0, weight=0) # "Urządzenie:" label fixed width
        self.device_selection_frame.grid_columnconfigure(1, weight=1) # Combobox column expands
        self.device_selection_frame.grid_columnconfigure(2, weight=0) # "Edytuj Nazwę" button fixed width
        self.device_selection_frame.grid_rowconfigure(0, weight=0) # Top row (label, combobox, button) fixed height
        self.device_selection_frame.grid_rowconfigure(1, weight=1) # Status row expands vertically


        self.available_devices = []
        self.current_selected_device_id = None

        self.populate_devices() # Initial population of devices
        print("DEBUG: UsbDeviceControllerApp __init__ finished.") # DEBUG PRINT

    def populate_devices(self):
        """
        Retrieves and displays the list of available input devices in the combobox.
        Attempts to re-select the previously chosen device if it's still available.
        """
        old_selected_id = self.current_selected_device_id
        # Start a new thread to populate devices asynchronously, passing the device_db
        populate_thread = threading.Thread(target=self._perform_populate_devices_async, args=(old_selected_id, self.device_db))
        populate_thread.daemon = True
        populate_thread.start()

    def _perform_populate_devices_async(self, old_selected_id, device_database):
        """
        Performs device population in a separate thread.
        Schedules a GUI update in the main thread after completion.
        """
        self.available_devices = get_input_devices_by_instance_id_pattern(device_database) # Pass database
        self.master.after(0, self._update_populate_gui, old_selected_id)

    def _update_populate_gui(self, old_selected_id):
        """
        Updates the GUI elements related to device population.
        This method is called from the main Tkinter thread.
        """
        device_display_names = [dev["display_name"] for dev in self.available_devices]
        self.device_combobox['values'] = device_display_names

        if old_selected_id and any(dev["id"] == old_selected_id for dev in self.available_devices):
            selected_dev = next((dev for dev in self.available_devices if dev["id"] == old_selected_id), None)
            if selected_dev: # Ensure selected_dev is not None
                self.device_combobox.set(selected_dev["display_name"])
                self.current_selected_device_id = selected_dev["id"]
                self.update_selected_device_status()
                self.enable_buttons()
            else: # If old_selected_id is no longer valid, treat as no device selected
                self.device_combobox.set("Brak urządzeń wejściowych do wyboru")
                self.current_selected_device_id = None
                self.status_label_text.set("Status: Nie wybrano")
                self.status_indicator.config(fg="gray")
                self.disable_buttons()
        elif self.available_devices:
            self.device_combobox.set(self.available_devices[0]["display_name"])
            self.current_selected_device_id = self.available_devices[0]["id"]
            self.update_selected_device_status()
            self.enable_buttons()
        else:
            self.device_combobox.set("Brak urządzeń wejściowych do wyboru")
            self.current_selected_device_id = None
            self.status_label_text.set("Status: Nie wybrano")
            self.status_indicator.config(fg="gray")
            self.disable_buttons()

    def on_device_selected(self, event=None):
        """
        Callback function executed when a new device is selected in the combobox.
        Updates the selected device ID and its status in the UI.
        """
        selected_name = self.selected_device_name.get()
        # Find the ID corresponding to the selected display name
        selected_device_obj = next((dev for dev in self.available_devices if dev["display_name"] == selected_name), None)
        if selected_device_obj:
            self.current_selected_device_id = selected_device_obj["id"]
        else:
            self.current_selected_device_id = None # Should not happen if item is truly selected

        self.update_selected_device_status()
        # Buttons are enabled/disabled within update_selected_device_status/_update_status_gui

    def update_selected_device_status(self):
        """
        Initiates an asynchronous update of the selected device's status.
        The actual status check is done in a separate thread to keep the GUI responsive.
        """
        if self.current_selected_device_id:
            # Set a temporary status indicating update in progress
            self.status_label_text.set("Status: Odświeżanie...")
            self.status_indicator.config(fg="orange")
            self.disable_buttons() # Disable buttons during update

            # Start a new thread to perform the status check
            status_thread = threading.Thread(target=self._perform_status_check, args=(self.current_selected_device_id,))
            status_thread.daemon = True # Allow the main program to exit even if thread is running
            status_thread.start()
        else:
            self.status_indicator.config(fg="gray")
            self.status_label_text.set("Status: Nie wybrano")
            self.disable_buttons()

    def _perform_status_check(self, device_id):
        """
        Performs the device status check in a separate thread.
        Schedules a GUI update in the main thread after completion.
        """
        status = get_device_status(device_id)
        # Schedule the GUI update to run in the main Tkinter thread
        self.master.after(0, self._update_status_gui, status)

    def _update_status_gui(self, status):
        """
        Updates the GUI elements with the fetched device status.
        This method is called from the main Tkinter thread.
        """
        if status == "enabled":
            self.status_indicator.config(fg="green")
            self.status_label_text.set("Status: Włączone")
        elif status == "disabled":
            self.status_indicator.config(fg="red")
            self.status_label_text.set("Status: Wyłączone")
        else:
            self.status_indicator.config(fg="gray")
            self.status_label_text.set("Status: Nieznany / Problem")
        self.enable_buttons() # Re-enable buttons after update

    def enable_buttons(self):
        """Enables all action buttons."""
        self.disable_button.config(state="normal")
        self.enable_button.config(state="normal")
        # Ensure edit_name_button is enabled here
        self.edit_name_button.config(state="normal") 

    def disable_buttons(self):
        """Disables all action buttons."""
        self.disable_button.config(state="disabled")
        self.enable_button.config(state="disabled")
        # Ensure edit_name_button is disabled here
        self.edit_name_button.config(state="disabled") 

    def toggle_device_status(self, device_id, action):
        """
        Toggles the status of a device (enable/disable) using pnputil.exe.
        Displays success or error messages to the user.
        This is now a method of the class.
        """
        command = ["pnputil", f"/{action}-device", device_id]
        stdout, stderr, returncode = run_as_admin(command, use_powershell=False)

        if returncode == 0:
            message = f"Pomyślnie {action}owano urządzenie:\n{device_id}"
            if stdout:
                message += f"\n\nSzczegóły:\n{stdout.strip()}"
            messagebox.showinfo("Sukces", message)
        else:
            error_message = f"Wystąpił błąd podczas {action}owania urządzenia:\n{device_id}"
            error_message += f"\n\nKod błędu: {returncode}"
            if stdout:
                error_message += f"\nSTDOUT: {stdout.strip()}"
            if stderr:
                error_message += f"\nSTDERR: {stderr.strip()}"
            messagebox.showerror("Błąd Operacji", error_message)

        # After toggling, update the UI status
        self.update_selected_device_status()

    def disable_selected_device(self):
        """Initiates the disable operation for the currently selected device."""
        if self.current_selected_device_id:
            self.toggle_device_status(self.current_selected_device_id, "disable")
        else:
            messagebox.showwarning("Ostrzeżenie", "Nie wybrano urządzenia do wyłączenia.")

    def enable_selected_device(self):
        """Initiates the enable operation for the currently selected device."""
        if self.current_selected_device_id:
            self.toggle_device_status(self.current_selected_device_id, "enable")
        else:
            messagebox.showwarning("Ostrzeżenie", "Nie wybrano urządzenia do włączenia.")

    def _open_edit_name_dialog(self):
        """
        Opens a modal dialog to allow the user to edit the custom name of the selected device.
        """
        # Get the currently selected display name from the combobox
        selected_display_name = self.selected_device_name.get()

        if not selected_display_name or selected_display_name == "Brak urządzeń wejściowych do wyboru":
            messagebox.showwarning("Błąd", "Najpierw wybierz urządzenie do edycji.")
            return

        # Find the actual device object using the display name (which includes VID:PID if no custom name)
        selected_device = next((dev for dev in self.available_devices if dev["display_name"] == selected_display_name), None)

        if not selected_device or not selected_device.get("vid_pid_key"): # Use .get() for safety
            messagebox.showwarning("Błąd", "Nie można edytować nazwy tego urządzenia (brak klucza VID:PID lub urządzenie niezidentyfikowane).")
            return

        device_key = selected_device["vid_pid_key"]
        current_custom_name = self.device_db.get(device_key, "") # Get existing custom name or empty string

        dialog = tk.Toplevel(self.master)
        dialog.title(f"Edytuj nazwę: {selected_device['display_name']}")
        dialog.transient(self.master) # Make it transient to the main window
        dialog.grab_set() # Make it modal
        dialog.geometry("450x200") # Slightly larger for better spacing
        center_window(dialog)
        dialog.resizable(False, False)

        # Label showing the device being edited
        tk.Label(dialog, text=f"Urządzenie: {selected_device['display_name']}", wraplength=400, justify="left", font=("Arial", 10, "bold")).pack(pady=5)
        
        # Current custom name (informational)
        tk.Label(dialog, text=f"Obecna nazwa niestandardowa: {current_custom_name if current_custom_name else 'Brak'}", font=("Arial", 9, "italic")).pack(pady=2)

        # --- Updated instruction text for clarity ---
        tk.Label(dialog, text="Wpisz nową nazwę (pozostaw puste, aby przywrócić domyślną nazwę systemową):", wraplength=400, justify="left").pack(pady=5)

        new_name_var = tk.StringVar(value=current_custom_name)
        name_entry = ttk.Entry(dialog, textvariable=new_name_var, width=60, font=("Arial", 10))
        name_entry.pack(pady=5)
        name_entry.focus_set()

        def save_and_close():
            new_name = new_name_var.get().strip()
            self._save_custom_name(new_name, device_key)
            dialog.destroy()

        def cancel_and_close():
            dialog.destroy()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)

        save_button = tk.Button(button_frame, text="Zapisz", command=save_and_close, bg="#66FF66", fg="black", font=("Arial", 10, "bold"))
        save_button.pack(side=tk.LEFT, padx=10, pady=5)

        cancel_button = tk.Button(button_frame, text="Anuluj", command=cancel_and_close, bg="#FF6666", fg="white", font=("Arial", 10, "bold"))
        cancel_button.pack(side=tk.RIGHT, padx=10, pady=5)

        self.master.wait_window(dialog) # Wait for dialog to close

    def _save_custom_name(self, new_name, device_key):
        """
        Saves or removes a custom name for a device in the database.
        """
        if new_name:
            self.device_db[device_key] = new_name
            messagebox.showinfo("Sukces", f"Nazwa urządzenia dla {device_key} została zaktualizowana na: '{new_name}'.")
        elif device_key in self.device_db: # If new_name is empty, remove existing custom name
            del self.device_db[device_key]
            # --- Changed message for clarity ---
            messagebox.showinfo("Sukces", f"Niestandardowa nazwa urządzenia dla {device_key} została usunięta, przywrócono nazwę domyślną.")
        else:
            # No custom name to save or delete
            return

        save_device_database(self.device_db) # Save the updated database to file
        self.populate_devices() # Refresh the device list in the main window

def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

if __name__ == "__main__":
    print("DEBUG: Script started.") # DEBUG PRINT
    # Ensure the script runs with administrator privileges from the start.
    # The run_as_admin function handles re-launching if needed.
    # It passes dummy arguments to allow the privilege check to happen.
    if not ctypes.windll.shell32.IsUserAnAdmin(): # Corrected from IsUserAnAdmin
        print("DEBUG: Not running as admin. Attempting to re-launch with admin privileges.") # DEBUG PRINT
        # Ensure __file__ path is correctly quoted to handle spaces.
        script_path = os.path.abspath(sys.argv[0]) # Get path to current script
        quoted_script_path = f'"{script_path}"' # Ensure proper quoting for paths with spaces

        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, quoted_script_path, None, 1)
        sys.exit(0) # Exit the current non-elevated process
    else:
        print("DEBUG: Running as admin. Proceeding with GUI initialization.") # DEBUG PRINT
        # The global exception handler will catch any errors here or later.
        root = tk.Tk()
        app = UsbDeviceControllerApp(root)
        center_window(root)
        print("DEBUG: Calling root.mainloop().") # DEBUG PRINT
        root.mainloop()
        print("DEBUG: root.mainloop() finished.") # DEBUG PRINT
    print("DEBUG: Script exited.") # DEBUG PRINT
