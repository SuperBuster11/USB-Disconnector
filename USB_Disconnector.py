import os
import subprocess
import ctypes
import sys
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import threading # Import for multi-threading
import tempfile # For temporary file creation
import json # Import for JSON parsing


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
    if ctypes.windll.shell32.IsUserAnAdmin():
        temp_file_path = None # Initialize to None for cleanup in finally block

        try:
            encoding_to_use = 'cp1250' # Standard encoding for Windows console output in many regions

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
                # '-Encoding Default' matches system's default encoding (usually cp1252 on Windows).
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
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{__file__}"', None, 1)
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
        if stderr:
            error_message += f"\nSTDERR: {stderr.strip()}"
        messagebox.showerror("Błąd Operacji", error_message)

    # After toggling, update the UI status if the app instance exists
    if UsbDeviceControllerApp.instance:
        # Call update_selected_device_status to refresh UI after toggle
        UsbDeviceControllerApp.instance.update_selected_device_status()


def get_input_devices_by_instance_id_pattern():
    """
    Retrieves a list of USB input devices (HID, Keyboard, Mouse, PointingDevice, GamePort)
    using PowerShell's Get-PnpDevice and converts the output to JSON for robust parsing.
    """
    ps_command = (
        r"Get-PnpDevice -PresentOnly | " # Only get devices currently connected
        r"Where-Object { ($_.InstanceId -like 'USB\VID_*&PID_*\*') -and " # Filter for USB devices with VID/PID pattern
        r"($_.Class -eq 'HIDClass' -or $_.Class -eq 'Keyboard' -or $_.Class -eq 'Mouse' -or $_.Class -eq 'PointingDevice' -or $_.Class -eq 'GamePort') } | " # Filter by specific device classes
        r"Select-Object DeviceDescription, FriendlyName, InstanceId, HardwareId | ConvertTo-Json -Compress" # Select and format as JSON
    )
    
    stdout, stderr, returncode = run_as_admin(ps_command, use_powershell=True)

    devices_list = []
    if returncode == 0 and stdout:
        try:
            json_data = json.loads(stdout)
            # Ensure json_data is a list, even if ConvertTo-Json returns a single object
            if not isinstance(json_data, list):
                json_data = [json_data] # Wrap single object in a list

            for dev_info in json_data:
                # Extract relevant info from dictionary using .get() for safety
                display_name = dev_info.get("FriendlyName") or dev_info.get("DeviceDescription") or dev_info.get("InstanceId")
                hardware_id_list = dev_info.get("HardwareId") # HardwareId might be an array in JSON or a string
                hardware_id = ""
                if hardware_id_list:
                    if isinstance(hardware_id_list, list) and hardware_id_list: # If it's a non-empty list
                        hardware_id = hardware_id_list[0].strip()
                    elif isinstance(hardware_id_list, str): # If it's a string
                        hardware_id = hardware_id_list.strip()
                    
                    if hardware_id and len(hardware_id) > 5 and hardware_id not in display_name:
                        display_name = f"{display_name} ({hardware_id})"

                # InstanceId is crucial, ensure it's present before adding to list
                if dev_info.get("InstanceId"):
                    devices_list.append({
                        "display_name": display_name,
                        "id": dev_info["InstanceId"],
                        "hardware_id": hardware_id,
                        "full_info": dev_info # Store full info as a dictionary
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
        print(f"  STDERR (raw): '{stderr}'")
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
        UsbDeviceControllerApp.instance = self # Assign the current instance to the class variable
        self.master = master
        master.title("Kontroler Urządzeń Wejściowych USB")
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
        # --- End Add window icon ---

        # Frame for device selection
        self.device_selection_frame = ttk.LabelFrame(master, text="Wybierz urządzenie wejściowe", padding=(10, 10))
        self.device_selection_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Device selection combobox
        tk.Label(self.device_selection_frame, text="Urządzenie:", font=("Arial", 11)).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.selected_device_name = tk.StringVar()
        self.device_combobox = ttk.Combobox(self.device_selection_frame, textvariable=self.selected_device_name, width=80, state="readonly", font=("Arial", 10))
        self.device_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew", columnspan=3)
        self.device_combobox.bind("<<ComboboxSelected>>", self.on_device_selected)

        # Status indicator (colored circle) and text label
        self.status_indicator = tk.Label(self.device_selection_frame, text="●", font=("Arial", 18, "bold"), width=2)
        self.status_indicator.grid(row=1, column=0, padx=5, pady=10, sticky="w")
        self.status_label_text = tk.StringVar(value="Status: Nieznany")
        self.status_text_label = tk.Label(self.device_selection_frame, textvariable=self.status_label_text, font=("Arial", 11))
        self.status_text_label.grid(row=1, column=1, padx=5, pady=10, sticky="w", columnspan=3)

        # Action buttons
        self.disable_button = tk.Button(master, text="Wyłącz wybrane urządzenie",
                                         command=self.disable_selected_device,
                                         bg="#FF6666", fg="white", width=25, height=2, font=("Arial", 10, "bold"), relief="raised", state="disabled")
        self.disable_button.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.enable_button = tk.Button(master, text="Włącz wybrane urządzenie",
                                        command=self.enable_selected_device,
                                        bg="#66FF66", fg="black", width=25, height=2, font=("Arial", 10, "bold"), relief="raised", state="disabled")
        self.enable_button.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        self.refresh_button = tk.Button(master, text="Odśwież listę urządzeń",
                                         command=self.populate_devices,
                                         bg="#ADD8E6", fg="black", width=25, height=2, font=("Arial", 10, "bold"), relief="raised")
        self.refresh_button.grid(row=3, column=0, padx=10, pady=5, sticky="ew")

        # Admin privilege reminder
        tk.Label(master, text="Pamiętaj: program wymaga uprawnień administratora.", font=("Arial", 8, "italic"), fg="gray") \
            .grid(row=4, column=0, pady=(10, 5))

        # Configure grid weights for responsiveness (though window is not resizable)
        master.grid_rowconfigure(0, weight=1)
        master.grid_columnconfigure(0, weight=1)
        self.device_selection_frame.grid_columnconfigure(1, weight=1)
        self.device_selection_frame.grid_rowconfigure(1, weight=1)

        self.available_devices = []
        self.current_selected_device_id = None

        self.populate_devices() # Initial population of devices

    def populate_devices(self):
        """
        Retrieves and displays the list of available input devices in the combobox.
        Attempts to re-select the previously chosen device if it's still available.
        """
        old_selected_id = self.current_selected_device_id
        # Start a new thread to populate devices asynchronously
        populate_thread = threading.Thread(target=self._perform_populate_devices_async, args=(old_selected_id,))
        populate_thread.daemon = True
        populate_thread.start()

    def _perform_populate_devices_async(self, old_selected_id):
        """
        Performs device population in a separate thread.
        Schedules a GUI update in the main thread after completion.
        """
        self.available_devices = get_input_devices_by_instance_id_pattern()
        self.master.after(0, self._update_populate_gui, old_selected_id)

    def _update_populate_gui(self, old_selected_id):
        """
        Updates the GUI elements related to device population.
        This method is called from the main Tkinter thread.
        """
        device_display_names = [dev["display_name"] for dev in self.available_devices]
        self.device_combobox['values'] = device_display_names

        if old_selected_id and any(dev["id"] == old_selected_id for dev in self.available_devices):
            selected_dev = next(dev for dev in self.available_devices if dev["display_name"] == self.selected_device_name.get())
            self.device_combobox.set(selected_dev["display_name"])
            self.current_selected_device_id = selected_dev["id"]
            self.update_selected_device_status()
            self.enable_buttons()
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
        self.current_selected_device_id = next((dev["id"] for dev in self.available_devices if dev["display_name"] == selected_name), None)
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
        """Enables the disable and enable action buttons."""
        self.disable_button.config(state="normal")
        self.enable_button.config(state="normal")

    def disable_buttons(self):
        """Disables the disable and enable action buttons."""
        self.disable_button.config(state="disabled")
        self.enable_button.config(state="disabled")

    def disable_selected_device(self):
        """Initiates the disable operation for the currently selected device."""
        if self.current_selected_device_id:
            toggle_device_status(self.current_selected_device_id, "disable")
        else:
            messagebox.showwarning("Ostrzeżenie", "Nie wybrano urządzenia do wyłączenia.")

    def enable_selected_device(self):
        """Initiates the enable operation for the currently selected device."""
        if self.current_selected_device_id:
            toggle_device_status(self.current_selected_device_id, "enable")
        else:
            messagebox.showwarning("Ostrzeżenie", "Nie wybrano urządzenia do włączenia.")


if __name__ == "__main__":
    # Ensure the script runs with administrator privileges from the start.
    # The run_as_admin function handles re-launching if needed.
    # It passes dummy arguments to allow the privilege check to happen.
    if not ctypes.windll.shell32.IsUserAnAdmin():
        # This will attempt to re-launch the script as admin and exit the current process
        run_as_admin([]) # Pass an empty list as arguments for the initial check

    # If we reach here, the script is running with admin privileges
    root = tk.Tk()
    app = UsbDeviceControllerApp(root)
    root.mainloop()

