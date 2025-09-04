import tkinter as tk
from tkinter import ttk, messagebox
import serial
import serial.tools.list_ports
import csv
import os
import threading
import time
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import numpy as np

# Origin COM interface
try:
    import win32com.client
    import pythoncom
    ORIGIN_AVAILABLE = True
except ImportError:
    ORIGIN_AVAILABLE = False

# Localization
from localization import _, localization

class TemperatureMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title(_("app_title"))
        self.root.geometry("800x600")

        # Initialize variables
        self.serial_connection = None
        self.is_measuring = False
        self.csv_file = None
        self.csv_writer = None
        self.measurement_start_time = None
        self.previous_temperature = None
        self.previous_time = None

        # Origin COM variables
        self.origin_app = None
        self.origin_worksheet = None
        self.origin_enabled = False
        self.origin_worksheet_name = None
        self.origin_current_row = 1  # Track current row for data writing

        # Data for plotting
        self.timestamps = []
        self.temperatures = []
        self.changes = []

        # Create measurements folder
        config_folder = localization.get_config('Files', 'folder')
        if config_folder:
            self.measurements_folder = config_folder
        else:
            documents_path = os.path.expanduser("~/Documents")
            self.measurements_folder = os.path.join(documents_path, "TemperatureMonitor")

        if not os.path.exists(self.measurements_folder):
            os.makedirs(self.measurements_folder)

        self.setup_gui()
        self.setup_plot()

    def setup_gui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text=_("main_title"),
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Control buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10))

        # Start button
        self.start_button = ttk.Button(button_frame, text=_("start_button"),
                                      command=self.start_measurement)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))

        # Stop button
        self.stop_button = ttk.Button(button_frame, text=_("stop_button"),
                                     command=self.stop_measurement, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(0, 20))

        # Origin checkbox
        self.origin_var = tk.BooleanVar()
        self.origin_checkbox = ttk.Checkbutton(button_frame, text=_("origin_checkbox"),
                                             variable=self.origin_var, command=self.on_origin_checkbox_change)
        self.origin_checkbox.pack(side=tk.LEFT, padx=(0, 20))

        # Browse Measurements button
        self.browse_button = ttk.Button(button_frame, text=_("browse_button"),
                                       command=self.browse_measurements)
        self.browse_button.pack(side=tk.LEFT)

        # Status label
        self.status_label = ttk.Label(main_frame, text=_("ready_status"),
                                     foreground="blue")
        self.status_label.grid(row=2, column=0, columnspan=3, pady=(0, 10))

        # Plot frame
        self.plot_frame = ttk.Frame(main_frame)
        self.plot_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.plot_frame.columnconfigure(0, weight=1)
        self.plot_frame.rowconfigure(0, weight=1)

    def setup_plot(self):
        # Create matplotlib figure
        self.fig = Figure(figsize=(8, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_title(_("plot_title"))
        self.ax.set_xlabel(_("plot_xlabel"))
        self.ax.set_ylabel(_("plot_ylabel"))
        self.ax.grid(True, alpha=0.3)

        # Create canvas
        self.canvas = FigureCanvasTkAgg(self.fig, self.plot_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    def find_arduino_port(self):
        """Find the Arduino port automatically"""
        ports = serial.tools.list_ports.comports()
        for port in ports:
            # Common Arduino identifiers
            if any(identifier in port.description.lower() for identifier in
                   ['ch340']):
                return port.device
        return None

    def on_origin_checkbox_change(self):
        """Handle Origin checkbox state change"""
        self.origin_enabled = self.origin_var.get()
        if self.origin_enabled:
            self.update_status(_("origin_enabled"), "blue")
        else:
            self.update_status(_("origin_disabled"), "blue")

    def browse_measurements(self):
        """Open the measurements folder in Windows Explorer"""
        try:
            # Ensure the measurements folder exists
            if not os.path.exists(self.measurements_folder):
                os.makedirs(self.measurements_folder)
                self.update_status(_("folder_created", path=self.measurements_folder), "blue")

            # Open the folder
            os.startfile(self.measurements_folder)  # Windows
            self.update_status(_("folder_opened"), "green")

        except Exception as e:
            self.update_status(_("folder_error", error=str(e)), "red")

    def connect_to_origin(self):
        """Connect to Origin application via COM"""
        if not ORIGIN_AVAILABLE:
            return False, _("origin_pywin32_error")

        try:
            # Try different COM interfaces for Origin 2018 compatibility
            try:
                # Method 1: Try Origin.ApplicationSI (Origin 2018 compatible)
                self.origin_app = win32com.client.Dispatch("Origin.ApplicationSI")
                return True, _("origin_connected")
            except:
                try:
                    # Method 2: Try to connect to existing Origin instance
                    self.origin_app = win32com.client.GetActiveObject("Origin.Application")
                    return True, _("origin_connected")
                except:
                    try:
                        # Method 3: Try to create new Origin instance
                        self.origin_app = win32com.client.Dispatch("Origin.Application")
                        return True, _("origin_created")
                    except:
                        # Method 4: Try Origin.ApplicationSI with GetActiveObject
                        self.origin_app = win32com.client.GetActiveObject("Origin.ApplicationSI")
                        return True, _("origin_connected")
        except Exception as e:
            return False, _("origin_connect_error", error=str(e))

    def create_origin_worksheet(self):
        """Create a new worksheet in Origin"""
        try:
            # Generate worksheet name with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.origin_worksheet_name = f"Measurement_{timestamp}"
            self.origin_worksheet = None

            # Try different methods to create worksheet for Origin 2018 compatibility
            methods = [
                self._create_worksheet_method1,
                self._create_worksheet_method2,
                self._create_worksheet_method3,
                self._create_worksheet_method4,
                self._create_worksheet_method5
            ]

            for method in methods:
                try:
                    if method():
                        break
                except Exception as e:
                    continue

            # Check if worksheet was created successfully
            if self.origin_worksheet is None:
                return False, _("origin_worksheet_error", error="All worksheet creation methods failed")

            # Reset row counter for new worksheet
            self.origin_current_row = 1

            # Column headers will be set automatically when data is written

            # Worksheet created successfully

            return True, _("origin_worksheet_created", name=self.origin_worksheet_name)
        except Exception as e:
            return False, _("origin_worksheet_error", error=str(e))

    def _create_worksheet_method1(self):
        """Method 1: Try Worksheets.Add (newer versions)"""
        try:
            self.origin_worksheet = self.origin_app.Worksheets.Add(self.origin_worksheet_name)
            return self.origin_worksheet is not None
        except:
            return False

    def _create_worksheet_method2(self):
        """Method 2: Try CreatePage method to create a workbook (Origin 2018 compatible)"""
        try:
            # CreatePage(type, name, template) where type 2 = workbook
            workbook_name = self.origin_app.CreatePage(2, f"Book_{self.origin_worksheet_name}", "origin")
            if workbook_name:
                # Access the first worksheet using WorksheetPages and Layers
                self.origin_worksheet = self.origin_app.WorksheetPages(workbook_name).Layers(0)
                if self.origin_worksheet:
                    # Try to rename the worksheet
                    try:
                        self.origin_worksheet.Name = self.origin_worksheet_name
                    except:
                        pass  # Renaming failed, but worksheet exists
                    return True
            return False
        except:
            return False

    def _create_worksheet_method3(self):
        """Method 3: Try creating a worksheet directly (type 1)"""
        try:
            page_name = self.origin_app.CreatePage(1, self.origin_worksheet_name, "origin")
            if page_name:
                # Access the worksheet using WorksheetPages
                self.origin_worksheet = self.origin_app.WorksheetPages(page_name).Layers(0)
                return self.origin_worksheet is not None
            return False
        except:
            return False

    def _create_worksheet_method4(self):
        """Method 4: Try using the first available worksheet"""
        try:
            # Get all worksheet pages
            worksheet_pages = self.origin_app.WorksheetPages
            if worksheet_pages and worksheet_pages.Count > 0:
                # Get the first worksheet page
                first_page = worksheet_pages(0)
                if first_page:
                    # Get the first layer (worksheet) from that page
                    self.origin_worksheet = first_page.Layers(0)
                    if self.origin_worksheet:
                        # Try to rename it
                        try:
                            self.origin_worksheet.Name = self.origin_worksheet_name
                        except:
                            pass  # Renaming failed, but worksheet exists
                        return True
            return False
        except:
            return False

    def _create_worksheet_method5(self):
        """Method 5: Create a simple workbook and access it"""
        try:
            workbook_name = self.origin_app.CreatePage(2, "TempBook", "origin")
            if workbook_name:
                self.origin_worksheet = self.origin_app.WorksheetPages(workbook_name).Layers(0)
                if self.origin_worksheet:
                    # Try to rename it
                    try:
                        self.origin_worksheet.Name = self.origin_worksheet_name
                    except:
                        pass  # Renaming failed, but worksheet exists
                    return True
            return False
        except:
            return False

    def _write_to_origin_main_thread(self, elapsed_ms, temperature, change):
        """Write data to Origin from main thread (COM threading fix)"""
        try:
            success, message = self.write_to_origin(elapsed_ms, temperature, change)
            if not success:
                self.update_status(f"Origin Error: {message}", "red")
            else:
                # Every 10 data points, check what's in the worksheet
                if self.origin_current_row % 10 == 0:
                    self._debug_worksheet_contents()
        except Exception as e:
            self.update_status(f"Origin Error: {str(e)}", "red")

    def _debug_worksheet_contents(self):
        """Debug method to check worksheet contents"""
        # Debug method removed - data writing is working correctly
        pass

    def write_to_origin(self, elapsed_ms, temperature, change):
        """Write data point to Origin worksheet"""
        try:
            if self.origin_worksheet is None:
                return False, _("origin_no_worksheet")

            # Initialize COM for this thread
            try:
                pythoncom.CoInitialize()
            except:
                pass  # Already initialized

            # Try the working LabTalk method first, then fallback methods
            methods = [
                ("LabTalk_Execute", lambda: self._write_to_origin_method17(elapsed_ms, temperature, change)),
                ("Cols_Method", lambda: self._write_to_origin_method18(elapsed_ms, temperature, change)),
                ("NewDataRange", lambda: self._write_to_origin_method19(elapsed_ms, temperature, change)),
                ("SetBinaryStorageData", lambda: self._write_to_origin_method16(elapsed_ms, temperature, change)),
                ("Range_Method", lambda: self._write_to_origin_method15(elapsed_ms, temperature, change)),
                ("SetData_Individual", lambda: self._write_to_origin_method13(elapsed_ms, temperature, change)),
                ("SetData_Array", lambda: self._write_to_origin_method14(elapsed_ms, temperature, change)),
                ("Simple_Cell_Assignment", lambda: self._write_to_origin_method11(elapsed_ms, temperature, change)),
                ("PutWorksheet_FullPath", lambda: self._write_to_origin_method12(elapsed_ms, temperature, change)),
                ("PutWorksheet_Worksheet", lambda: self._write_to_origin_method1(elapsed_ms, temperature, change)),
                ("PutWorksheet_App", lambda: self._write_to_origin_method2(elapsed_ms, temperature, change)),
                ("RowCounter_SetCell", lambda: self._write_to_origin_method9(elapsed_ms, temperature, change)),
                ("RowCounter_PutCell", lambda: self._write_to_origin_method10(elapsed_ms, temperature, change)),
                ("Cells_Property", lambda: self._write_to_origin_method3(elapsed_ms, temperature, change)),
                ("SetCell_Method", lambda: self._write_to_origin_method4(elapsed_ms, temperature, change)),
                ("Column_Access", lambda: self._write_to_origin_method5(elapsed_ms, temperature, change)),
                ("PutCell_Method", lambda: self._write_to_origin_method6(elapsed_ms, temperature, change)),
                ("SetCell_1Based", lambda: self._write_to_origin_method7(elapsed_ms, temperature, change)),
                ("Array_Assignment", lambda: self._write_to_origin_method8(elapsed_ms, temperature, change))
            ]

            for method_name, method in methods:
                try:
                    success, message = method()
                    if success:
                        return success, message
                except Exception as e:
                    continue

            return False, _("origin_write_error", error="All write methods failed")
        except Exception as e:
            return False, _("origin_write_error", error=str(e))
        finally:
            # Clean up COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def _write_to_origin_method1(self, elapsed_ms, temperature, change):
        """Method 1: Try PutWorksheet method on worksheet object (Origin 2018 compatible)"""
        try:
            # Prepare data as 2D array
            data = [[elapsed_ms, temperature, change]]
            # Call PutWorksheet on the worksheet object, not the application
            self.origin_worksheet.PutWorksheet(data)
            return True, _("origin_write_success")
        except:
            return False, "PutWorksheet method failed"

    def _write_to_origin_method2(self, elapsed_ms, temperature, change):
        """Method 2: Try PutWorksheet with application-level call"""
        try:
            # Use the correct format for worksheet reference
            worksheet_name = f"[{self.origin_worksheet_name}]Sheet1"
            data = [[elapsed_ms, temperature, change]]
            success = self.origin_app.PutWorksheet(worksheet_name, data)
            if success:
                return True, _("origin_write_success")
            return False, "PutWorksheet returned False"
        except:
            return False, "PutWorksheet method failed"

    def _write_to_origin_method3(self, elapsed_ms, temperature, change):
        """Method 3: Try using Cells property"""
        try:
            # Find next empty row
            last_row = self._find_next_empty_row()
            if last_row is None:
                last_row = 1

            # Write data using Cells property
            self.origin_worksheet.Cells(last_row, 0).Value = elapsed_ms
            self.origin_worksheet.Cells(last_row, 1).Value = temperature
            self.origin_worksheet.Cells(last_row, 2).Value = change
            return True, _("origin_write_success")
        except:
            return False, "Cells method failed"

    def _write_to_origin_method4(self, elapsed_ms, temperature, change):
        """Method 4: Try using SetCell method"""
        try:
            # Find next empty row
            last_row = self._find_next_empty_row()
            if last_row is None:
                last_row = 1

            # Write data using SetCell method
            self.origin_worksheet.SetCell(last_row, 0, elapsed_ms)
            self.origin_worksheet.SetCell(last_row, 1, temperature)
            self.origin_worksheet.SetCell(last_row, 2, change)
            return True, _("origin_write_success")
        except:
            return False, "SetCell method failed"

    def _write_to_origin_method5(self, elapsed_ms, temperature, change):
        """Method 5: Try using direct column access"""
        try:
            # Find next empty row
            last_row = self._find_next_empty_row()
            if last_row is None:
                last_row = 1

            # Write data using direct column access
            self.origin_worksheet.Columns(0).SetCell(last_row, elapsed_ms)
            self.origin_worksheet.Columns(1).SetCell(last_row, temperature)
            self.origin_worksheet.Columns(2).SetCell(last_row, change)
            return True, _("origin_write_success")
        except:
            return False, "Column access method failed"

    def _write_to_origin_method6(self, elapsed_ms, temperature, change):
        """Method 6: Try using PutCell method (Origin 2018 compatible)"""
        try:
            # Find next empty row
            last_row = self._find_next_empty_row()
            if last_row is None:
                last_row = 1

            # Write data using PutCell method
            self.origin_worksheet.PutCell(last_row, 0, elapsed_ms)
            self.origin_worksheet.PutCell(last_row, 1, temperature)
            self.origin_worksheet.PutCell(last_row, 2, change)
            return True, _("origin_write_success")
        except:
            return False, "PutCell method failed"

    def _write_to_origin_method7(self, elapsed_ms, temperature, change):
        """Method 7: Try using SetCell with proper indexing (Origin 2018 compatible)"""
        try:
            # Find next empty row
            last_row = self._find_next_empty_row()
            if last_row is None:
                last_row = 1

            # Write data using SetCell with 1-based indexing
            self.origin_worksheet.SetCell(last_row, 1, elapsed_ms)  # Column 1
            self.origin_worksheet.SetCell(last_row, 2, temperature)  # Column 2
            self.origin_worksheet.SetCell(last_row, 3, change)       # Column 3
            return True, _("origin_write_success")
        except:
            return False, "SetCell with 1-based indexing failed"

    def _write_to_origin_method8(self, elapsed_ms, temperature, change):
        """Method 8: Try using direct data array assignment"""
        try:
            # Find next empty row
            last_row = self._find_next_empty_row()
            if last_row is None:
                last_row = 1

            # Write data using direct array assignment
            data_array = [elapsed_ms, temperature, change]
            for col in range(3):
                self.origin_worksheet.SetCell(last_row, col + 1, data_array[col])
            return True, _("origin_write_success")
        except:
            return False, "Direct array assignment failed"

    def _write_to_origin_method9(self, elapsed_ms, temperature, change):
        """Method 9: Try using row counter for efficient writing"""
        try:
            # Use the tracked row counter
            current_row = self.origin_current_row

            # Write data using SetCell with 1-based indexing
            self.origin_worksheet.SetCell(current_row, 1, elapsed_ms)  # Column 1
            self.origin_worksheet.SetCell(current_row, 2, temperature)  # Column 2
            self.origin_worksheet.SetCell(current_row, 3, change)       # Column 3

            # Increment row counter for next write
            self.origin_current_row += 1
            return True, _("origin_write_success")
        except:
            return False, "Row counter method failed"

    def _write_to_origin_method10(self, elapsed_ms, temperature, change):
        """Method 10: Try using PutCell with row counter"""
        try:
            # Use the tracked row counter
            current_row = self.origin_current_row

            # Write data using PutCell with 0-based indexing
            self.origin_worksheet.PutCell(current_row, 0, elapsed_ms)  # Column 0
            self.origin_worksheet.PutCell(current_row, 1, temperature)  # Column 1
            self.origin_worksheet.PutCell(current_row, 2, change)       # Column 2

            # Increment row counter for next write
            self.origin_current_row += 1
            return True, _("origin_write_success")
        except:
            return False, "PutCell with row counter failed"

    def _write_to_origin_method11(self, elapsed_ms, temperature, change):
        """Method 11: Try the simplest possible approach - direct cell assignment"""
        try:
            # Use the tracked row counter
            current_row = self.origin_current_row

            # Try the simplest possible cell assignment
            self.origin_worksheet.Cells(current_row, 1).Value = elapsed_ms
            self.origin_worksheet.Cells(current_row, 2).Value = temperature
            self.origin_worksheet.Cells(current_row, 3).Value = change

            # Increment row counter for next write
            self.origin_current_row += 1
            return True, _("origin_write_success")
        except Exception as e:
            return False, f"Simple cell assignment failed: {e}"

    def _write_to_origin_method12(self, elapsed_ms, temperature, change):
        """Method 12: Try using the application's PutWorksheet with full path"""
        try:
            # Try using the full workbook path
            full_path = f"[{self.origin_worksheet_name}]Sheet1"
            data = [[elapsed_ms, temperature, change]]
            success = self.origin_app.PutWorksheet(full_path, data)
            if success:
                return True, _("origin_write_success")
            return False, f"PutWorksheet with full path returned {success}"
        except Exception as e:
            return False, f"PutWorksheet with full path failed: {e}"

    def _write_to_origin_method13(self, elapsed_ms, temperature, change):
        """Method 13: Try using SetData method (Origin specific)"""
        try:
            # Use the tracked row counter
            current_row = self.origin_current_row

            # Prepare data as a list for SetData
            data = [elapsed_ms, temperature, change]
            
            # Use SetData method with row and column specification
            # SetData(row, col, data) - 1-based indexing
            self.origin_worksheet.SetData(current_row, 1, elapsed_ms)
            self.origin_worksheet.SetData(current_row, 2, temperature)
            self.origin_worksheet.SetData(current_row, 3, change)

            # Data written successfully - no debug output needed

            # Increment row counter for next write
            self.origin_current_row += 1
            return True, _("origin_write_success")
        except Exception as e:
            return False, f"SetData method failed: {e}"

    def _write_to_origin_method14(self, elapsed_ms, temperature, change):
        """Method 14: Try using SetData with array"""
        try:
            # Use the tracked row counter
            current_row = self.origin_current_row

            # Prepare data as array for SetData
            data = [elapsed_ms, temperature, change]
            
            # Try setting all data at once
            self.origin_worksheet.SetData(current_row, 1, data)

            # Increment row counter for next write
            self.origin_current_row += 1
            return True, _("origin_write_success")
        except Exception as e:
            return False, f"SetData array method failed: {e}"

    def _write_to_origin_method15(self, elapsed_ms, temperature, change):
        """Method 15: Try using Range method"""
        try:
            # Use the tracked row counter
            current_row = self.origin_current_row

            # Try using Range method
            range_obj = self.origin_worksheet.Range(f"A{current_row}:C{current_row}")
            range_obj.Value = [elapsed_ms, temperature, change]

            # Increment row counter for next write
            self.origin_current_row += 1
            return True, _("origin_write_success")
        except Exception as e:
            return False, f"Range method failed: {e}"

    def _write_to_origin_method16(self, elapsed_ms, temperature, change):
        """Method 16: Try using SetBinaryStorageData method"""
        try:
            # Use the tracked row counter
            current_row = self.origin_current_row

            # Prepare data as binary storage
            data = [elapsed_ms, temperature, change]
            
            # Try using SetBinaryStorageData
            self.origin_worksheet.SetBinaryStorageData(current_row, 1, elapsed_ms)
            self.origin_worksheet.SetBinaryStorageData(current_row, 2, temperature)
            self.origin_worksheet.SetBinaryStorageData(current_row, 3, change)

            # Increment row counter for next write
            self.origin_current_row += 1
            return True, _("origin_write_success")
        except Exception as e:
            return False, f"SetBinaryStorageData method failed: {e}"

    def _write_to_origin_method17(self, elapsed_ms, temperature, change):
        """Method 17: Try using Execute method with LabTalk commands"""
        try:
            # Use the tracked row counter
            current_row = self.origin_current_row

            # Try using Execute method with LabTalk commands
            cmd = f"col(A)[{current_row}] = {elapsed_ms}; col(B)[{current_row}] = {temperature}; col(C)[{current_row}] = {change};"
            self.origin_worksheet.Execute(cmd)

            # Increment row counter for next write
            self.origin_current_row += 1
            return True, _("origin_write_success")
        except Exception as e:
            return False, f"Execute method failed: {e}"

    def _write_to_origin_method18(self, elapsed_ms, temperature, change):
        """Method 18: Try using Cols method"""
        try:
            # Use the tracked row counter
            current_row = self.origin_current_row

            # Try using Cols method
            self.origin_worksheet.Cols(1)[current_row] = elapsed_ms
            self.origin_worksheet.Cols(2)[current_row] = temperature
            self.origin_worksheet.Cols(3)[current_row] = change

            # Increment row counter for next write
            self.origin_current_row += 1
            return True, _("origin_write_success")
        except Exception as e:
            return False, f"Cols method failed: {e}"

    def _write_to_origin_method19(self, elapsed_ms, temperature, change):
        """Method 19: Try using NewDataRange method"""
        try:
            # Use the tracked row counter
            current_row = self.origin_current_row

            # Try using NewDataRange method
            data_range = self.origin_worksheet.NewDataRange()
            data_range.SetData(current_row, 1, elapsed_ms)
            data_range.SetData(current_row, 2, temperature)
            data_range.SetData(current_row, 3, change)

            # Increment row counter for next write
            self.origin_current_row += 1
            return True, _("origin_write_success")
        except Exception as e:
            return False, f"NewDataRange method failed: {e}"

    def _find_next_empty_row(self):
        """Find the next empty row in the worksheet"""
        try:
            # Since we're using a row counter, just return the current row
            # This avoids checking for empty rows which can create columns
            return self.origin_current_row
        except:
            # Fallback to row 1
            return 1

    def disconnect_from_origin(self):
        """Disconnect from Origin"""
        try:
            if self.origin_app:
                self.origin_app = None
            self.origin_worksheet = None
            self.origin_worksheet_name = None
            return True, _("origin_disconnected")
        except Exception as e:
            return False, _("origin_disconnect_error", error=str(e))

    def start_measurement(self):
        """Start temperature measurement"""
        try:
            # Check Origin integration if enabled
            if self.origin_enabled:
                success, message = self.connect_to_origin()
                if not success:
                    self.update_status(_("origin_error", message=message), "red")
                    return
                self.update_status(f"Origin: {message}", "green")

                success, message = self.create_origin_worksheet()
                if not success:
                    self.update_status(_("origin_error", message=message), "red")
                    return
                self.update_status(f"Origin: {message}", "green")

            # Find Arduino port
            port = self.find_arduino_port()
            if not port:
                self.update_status(_("arduino_not_found"), "red")
                return

            # Connect to Arduino
            self.serial_connection = serial.Serial(port, 9600, timeout=0.5)
            time.sleep(2)  # Wait for Arduino to initialize

            # Create CSV file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"measurement_{timestamp}.csv"
            filepath = os.path.join(self.measurements_folder, filename)

            self.csv_file = open(filepath, 'w', newline='')
            self.csv_writer = csv.writer(self.csv_file)

            # Initialize measurement variables
            self.is_measuring = True
            self.measurement_start_time = time.time()
            self.previous_temperature = None
            self.previous_time = None

            # Clear previous data
            self.timestamps.clear()
            self.temperatures.clear()
            self.changes.clear()

            # Update GUI
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.origin_checkbox.config(state=tk.DISABLED)  # Disable Origin checkbox during measurement
            self.update_status(_("connected_status", port=port, filename=filename), "green")

            # Start measurement thread
            self.measurement_thread = threading.Thread(target=self.measure_loop, daemon=True)
            self.measurement_thread.start()

        except Exception as e:
            self.update_status(_("start_error", error=str(e)), "red")
            if self.serial_connection:
                self.serial_connection.close()
                self.serial_connection = None

    def stop_measurement(self):
        """Stop temperature measurement"""
        self.is_measuring = False

        # Close serial connection
        if self.serial_connection:
            self.serial_connection.close()
            self.serial_connection = None

        # Close CSV file
        if self.csv_file:
            self.csv_file.close()
            self.csv_file = None

        # Disconnect from Origin if enabled
        if self.origin_enabled:
            success, message = self.disconnect_from_origin()
            if success:
                self.update_status(_("stopped_origin_status", message=message), "blue")
            else:
                self.update_status(_("stopped_origin_error_status", message=message), "red")
        else:
            self.update_status(_("stopped_status"), "blue")

        # Update GUI
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.origin_checkbox.config(state=tk.NORMAL)  # Re-enable Origin checkbox when stopped

    def measure_loop(self):
        """Main measurement loop running in separate thread"""
        while self.is_measuring:
            try:
                # Read temperature from Arduino
                if self.serial_connection and self.serial_connection.in_waiting > 0:
                    line = self.serial_connection.readline().decode('utf-8').strip()
                    if line:
                        temperature = float(line)
                        current_time = time.time()

                        # Calculate milliseconds elapsed since measurement start (excluding Arduino init delay)
                        elapsed_ms = int((current_time - self.measurement_start_time) * 1000)

                        # Calculate change (first derivative per second)
                        change = 0.0
                        if self.previous_temperature is not None and self.previous_time is not None:
                            time_diff = current_time - self.previous_time
                            if time_diff > 0:
                                change = (temperature - self.previous_temperature) / time_diff

                        # Write to CSV
                        if self.csv_writer:
                            self.csv_writer.writerow([elapsed_ms, temperature, change])
                            self.csv_file.flush()

                        # Write to Origin if enabled (schedule in main thread)
                        if self.origin_enabled and self.origin_worksheet:
                            self.root.after(0, lambda: self._write_to_origin_main_thread(elapsed_ms, temperature, change))

                        # Store data for plotting
                        elapsed_time = current_time - self.measurement_start_time
                        self.timestamps.append(elapsed_time)
                        self.temperatures.append(temperature)
                        self.changes.append(change)

                        # Update plot (limit to last 100 points for performance)
                        if len(self.timestamps) > 100:
                            self.timestamps = self.timestamps[-100:]
                            self.temperatures = self.temperatures[-100:]
                            self.changes = self.changes[-100:]

                        # Update plot in main thread
                        self.root.after(0, self.update_plot)

                        # Update status with current temperature
                        self.root.after(0, lambda: self.update_status(
                            _("recording_status", temperature=temperature, change=change), "green"))

                        # Store for next iteration
                        self.previous_temperature = temperature
                        self.previous_time = current_time

                time.sleep(0.1)  # Small delay to prevent excessive CPU usage

            except Exception as e:
                self.root.after(0, lambda: self.update_status(_("measurement_error", error=str(e)), "red"))
                break

    def update_plot(self):
        """Update the temperature plot"""
        if self.timestamps and self.temperatures:
            self.ax.clear()
            self.ax.plot(self.timestamps, self.temperatures, 'b-', linewidth=2, label=_("plot_legend"))
            self.ax.set_title(_("plot_title"))
            self.ax.set_xlabel(_("plot_xlabel"))
            self.ax.set_ylabel(_("plot_ylabel"))
            self.ax.grid(True, alpha=0.3)
            self.ax.legend()

            # Auto-scale axes
            if len(self.timestamps) > 1:
                self.ax.set_xlim(min(self.timestamps), max(self.timestamps))
                temp_min, temp_max = min(self.temperatures), max(self.temperatures)
                temp_range = temp_max - temp_min
                if temp_range > 0:
                    self.ax.set_ylim(temp_min - temp_range * 0.1, temp_max + temp_range * 0.1)

            self.canvas.draw()

    def update_status(self, message, color="black"):
        """Update the status label"""
        self.status_label.config(text=message, foreground=color)

    def on_closing(self):
        """Handle application closing"""
        if self.is_measuring:
            self.stop_measurement()

        # Disconnect from Origin if connected
        if self.origin_enabled:
            self.disconnect_from_origin()

        self.root.destroy()

def main():
    root = tk.Tk()
    app = TemperatureMonitor(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
