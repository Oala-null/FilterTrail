import sys
import threading
import os
import time
import json
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                             QWidget, QLabel, QMessageBox, QFileDialog, QHBoxLayout,
                             QProgressBar, QSplitter)
from PyQt5.QtCore import QUrl, Qt, QThread, pyqtSignal
from PyQt5.QtWebEngineWidgets import QWebEngineView

# Excel monitoring class - uses win32com to interface with Excel
class ExcelFilterMonitor:
    def __init__(self):
        self.excel = None
        self.workbook = None
        self.worksheet = None
        self.filter_history = []
        self.data_file = "filter_data.json"
        self.last_save_time = time.time()
        self.save_interval = 1  # Save every second
        self.stop_event = threading.Event()
        self.last_known_filters = {}  # Track previous filter state
        self.header_names = {}  # Cache for all column names in header row
        self.primary_key_column = 1  # Default to first column as primary key (1-based index)
        self.primary_key_name = ""   # Name of primary key column (will be populated on connect)
        
        # Create empty filter history file if it doesn't exist
        if not os.path.exists(self.data_file):
            with open(self.data_file, 'w') as f:
                json.dump([], f)
        
    def connect_to_excel(self):
        """Connect to running Excel instance and read all column headers"""
        try:
            # Import here to avoid PyInstaller issues
            import win32com.client
            import pythoncom
            
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            self.excel = win32com.client.GetActiveObject("Excel.Application")
            self.workbook = self.excel.ActiveWorkbook
            self.worksheet = self.excel.ActiveSheet
            
            # IMPORTANT: Read all header names immediately on connection
            self.read_all_headers()
            
            # Set primary key name based on the column
            self.primary_key_name = self.header_names.get(self.primary_key_column, f"Column {self.primary_key_column}")
            print(f"Primary key column: {self.primary_key_name} (column {self.primary_key_column})")
            
            return True
        except Exception as e:
            print(f"Error connecting to Excel: {e}")
            return False
            
    def read_all_headers(self):
        """Read all column headers and cache them for future use"""
        try:
            # Clear existing header names
            self.header_names = {}
            
            # Try to get header row
            if hasattr(self.worksheet, "UsedRange"):
                header_row = self.worksheet.UsedRange.Rows(1)
                
                # Get column count
                try:
                    columns_count = min(header_row.Columns.Count, 200)  # Cap at 200 columns
                except:
                    columns_count = 50  # Reasonable default
                
                # Read all headers at once
                try:
                    header_values = header_row.Value
                    if isinstance(header_values, tuple):
                        for col_idx, val in enumerate(header_values[0], 1):
                            if col_idx <= columns_count:
                                col_name = str(val) if val is not None else f"Column {col_idx}"
                                self.header_names[col_idx] = col_name
                except:
                    # Fallback to reading cell by cell if batch read fails
                    for i in range(1, columns_count + 1):
                        try:
                            cell_value = header_row.Cells(1, i).Value
                            col_name = str(cell_value) if cell_value is not None else f"Column {i}"
                            self.header_names[i] = col_name
                        except:
                            self.header_names[i] = f"Column {i}"
                
                print(f"Read {len(self.header_names)} column headers")
            else:
                print("Could not read headers - UsedRange not available")
        except Exception as e:
            print(f"Error reading headers: {e}")
    
    def background_monitoring(self, status_queue=None):
        """Run monitoring in background thread"""
        # Import here to avoid PyInstaller issues
        import win32com.client
        import pythoncom
        
        if status_queue:
            status_queue.put("Starting Excel monitor...")
            
        if not self.connect_to_excel():
            if status_queue:
                status_queue.put("Failed to connect to Excel")
            return False
            
        if status_queue:
            status_queue.put(f"Connected to Excel: {self.workbook.Name}, Sheet: {self.worksheet.Name}")
            
        print(f"Connected to Excel: {self.workbook.Name}, Sheet: {self.worksheet.Name}")
        print("Monitoring filter operations...")
        
        # Load existing filter history but with safeguards for performance
        try:
            if os.path.exists(self.data_file) and os.path.getsize(self.data_file) > 0:
                try:
                    with open(self.data_file, 'r') as f:
                        loaded_history = json.load(f)
                        
                    # Only keep a reasonable number of recent entries to avoid slowdowns
                    max_history_length = 100  # Adjust as needed for performance vs history length
                    if len(loaded_history) > max_history_length:
                        self.filter_history = loaded_history[-max_history_length:]
                        print(f"Loaded last {max_history_length} entries from history for performance")
                    else:
                        self.filter_history = loaded_history
                except Exception as e:
                    print(f"Error loading filter history from {self.data_file}: {e}")
                    # Try backup file if available
                    backup_file = "filter_data_backup.json"
                    if os.path.exists(backup_file):
                        try:
                            with open(backup_file, 'r') as f:
                                self.filter_history = json.load(f)
                            print(f"Loaded filter history from backup file")
                        except:
                            self.filter_history = []
                    else:
                        self.filter_history = []
        except Exception as e:
            print(f"Error accessing filter history file: {e}")
            self.filter_history = []
        
        last_filter_state = {}
        last_row_count = 0
        
        try:
            # Check if we already have filter history before adding initial state
            should_add_initial_state = True
            original_total_rows = 0
            
            # If we have prior filter history, use it instead of resetting to unfiltered state
            if self.filter_history:
                try:
                    # Get the original total rows from the first event
                    for event in self.filter_history:
                        if event.get("total_rows", 0) > 0:
                            original_total_rows = event.get("total_rows", 0)
                            break
                    
                    # Get the most recent state as our starting point
                    last_event = self.filter_history[-1]
                    last_row_count = last_event.get("current_row_count", 0)
                    
                    # Only add a new initial event if we're starting fresh
                    if original_total_rows > 0 and last_row_count > 0:
                        should_add_initial_state = False
                except:
                    # If anything fails, add a new initial state to be safe
                    should_add_initial_state = True
            
            # Create a timestamp for this session
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # IMPORTANT CHANGE: Get non-empty rows count based on primary key column instead of total rows
            try:
                # Get count of non-empty cells in primary key column
                primary_key_count = self.get_primary_key_count()
                print(f"Primary key column '{self.primary_key_name}' has {primary_key_count} non-empty cells")
                total_rows = primary_key_count  # Use this as our total row count
            except Exception as e:
                print(f"Error counting primary key cells: {e}")
                # Fallback to old method if primary key counting fails
                try:
                    # Direct Excel API count
                    total_rows = self.worksheet.UsedRange.Rows.Count - 1  # Subtract header row
                    if total_rows <= 0:  # Fallback if something went wrong
                        total_rows = self.worksheet.UsedRange.Rows.Count
                        if total_rows > 1:  # Allow for header row if more than 1 row
                            total_rows -= 1
                except:
                    # If direct count fails, try regular method
                    total_rows = self.get_total_row_count()
            
            print(f"Using {total_rows} rows based on primary key column")
            
            # If we have history and already know the original row count, use that
            if original_total_rows > 0:
                total_rows = original_total_rows
                print(f"Using original total rows from history: {total_rows}")
            
            # Only add initial state if we're starting fresh
            if should_add_initial_state:
                initial_event = {
                    "timestamp": timestamp,
                    "action": "initial_connection",
                    "added_filters": [],
                    "removed_filters": [],
                    "previous_row_count": total_rows,
                    "current_row_count": total_rows,
                    "total_rows": total_rows,
                    "filter_column": "All Data",
                    "filter_columns": []
                }
                self.filter_history.append(initial_event)
                self.save_filter_history()
            
            # Detect the display state - for accurate row counting
            display_state = {}
            try:
                display_state["DisplayPageBreaks"] = self.worksheet.DisplayPageBreaks
                display_state["DisplayGridlines"] = self.worksheet.DisplayGridlines
                display_state["DisplayHeadings"] = self.worksheet.DisplayHeadings
            except:
                pass
                
            # Immediately capture initial state and detect if filters are already applied
            try:
                # For the initial state, force a refresh of Excel view if possible
                try:
                    # Toggle a display setting to force refresh without affecting user
                    if "DisplayGridlines" in display_state:
                        current_setting = display_state["DisplayGridlines"]
                        self.worksheet.DisplayGridlines = not current_setting
                        time.sleep(0.1)
                        self.worksheet.DisplayGridlines = current_setting
                except:
                    pass
                
                # Get initial row count using multiple methods
                visible_rows = []
                for method in ['direct', 'special_cells', 'sampling']:
                    try:
                        if method == 'direct':
                            rows = self.get_direct_visible_count()
                        elif method == 'special_cells':
                            rows = self.get_special_cells_count()
                        else:
                            rows = self.get_sampling_count()
                        
                        if rows > 0:
                            visible_rows.append(rows)
                    except:
                        pass
                
                # Use the most accurate method based on results - prefer Excel's display
                if visible_rows:
                    # IMPORTANT FIX: We've found that Excel's tab display typically matches 
                    # the smallest count from our methods (not the largest)
                    # Using the smallest value gives more accurate results matching what Excel shows
                    min_count = min(visible_rows)
                    
                    # Add 5% validation - if smallest is much smaller than others, use median instead
                    if len(visible_rows) >= 3 and min_count < max(visible_rows) * 0.8:
                        # Use median for more reliability when there's big variance
                        visible_rows.sort()
                        last_row_count = visible_rows[len(visible_rows)//2]
                    else:
                        # Otherwise just use the minimum count as it best matches Excel's display
                        last_row_count = min_count
                else:
                    last_row_count = total_rows
                    
                print(f"Starting with {last_row_count} visible rows")
                    
                # Get initial filter state
                last_filter_state = self.get_current_filters()
                self.last_known_filters = dict(last_filter_state)  # Make a copy
                
                # Record initial state if filters are already active
                if last_filter_state and last_row_count < total_rows:
                    from datetime import datetime
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # Create filter event for initial state
                    added_filters = []
                    for col, filter_info in last_filter_state.items():
                        # Handle different formats of filter values
                        if isinstance(filter_info, dict) and "values" in filter_info:
                            filter_values = filter_info["values"]
                            column_index = filter_info.get("column_index", 0)
                        else:
                            filter_values = filter_info
                            column_index = 0
                            
                        added_filters.append({
                            "column": col, 
                            "values": filter_values,
                            "column_index": column_index
                        })
                    
                    # Collect all filter columns
                    filter_columns = [f["column"] for f in added_filters]
                    
                    # Use the first column as the display name if available, otherwise "Initial State"
                    display_filter_column = filter_columns[0] if filter_columns else "Initial State"
                    
                    initial_event = {
                        "timestamp": timestamp,
                        "action": "initial_state",
                        "added_filters": added_filters,
                        "removed_filters": [],
                        "previous_row_count": total_rows,
                        "current_row_count": last_row_count,
                        "total_rows": total_rows,
                        "filter_column": display_filter_column,     # Single column displayed
                        "filter_columns": filter_columns,           # All columns that changed in this step
                        "active_filters": list(last_filter_state.keys()) # All active filters
                    }
                    
                    print(f"Initial filter state detected with {len(added_filters)} active filters")
                    self.filter_history.append(initial_event)
                    self.save_filter_history()
            except Exception as e:
                print(f"Warning: Could not get initial filter state: {e}")
                last_filter_state = {}
            
            # Main monitoring loop
            while not self.stop_event.is_set():
                try:
                    # Skip operations if Excel becomes unresponsive
                    if not self.is_excel_alive():
                        print("Excel connection lost. Waiting for Excel...")
                        time.sleep(1)
                        if self.connect_to_excel():
                            print("Reconnected to Excel")
                        continue
                    
                    # Get current filter state    
                    current_filters = self.get_current_filters()
                except Exception as e:
                    time.sleep(0.5)
                    continue
                
                try:
                    # Get current row count using multiple methods for accuracy
                    visible_rows = []
                    for method in ['direct', 'special_cells', 'sampling']:
                        try:
                            if method == 'direct':
                                rows = self.get_direct_visible_count()
                            elif method == 'special_cells':
                                rows = self.get_special_cells_count()
                            else:
                                rows = self.get_sampling_count()
                            
                            if rows > 0:
                                visible_rows.append(rows)
                        except:
                            pass
                    
                    # Use the most accurate method based on results - match Excel's display
                    if visible_rows:
                        # IMPORTANT FIX: Excel's tab display typically matches 
                        # the smallest count from our methods (not the largest)
                        min_count = min(visible_rows)
                        
                        # Validation - if smallest is much smaller than others, use median instead
                        if len(visible_rows) >= 3 and min_count < max(visible_rows) * 0.8:
                            # Use median for more reliability when there's big variance
                            visible_rows.sort()
                            current_row_count = visible_rows[len(visible_rows)//2]
                        else:
                            # Use the minimum count as it best matches Excel's display
                            current_row_count = min_count
                    else:
                        # If no methods worked, try directly from Excel
                        current_row_count = self.get_visible_row_count()
                except Exception as e:
                    time.sleep(0.5)
                    continue
                
                # Check if anything changed
                row_count_changed = abs(current_row_count - last_row_count) > 2  # Reduce threshold to catch smaller changes
                filter_state_changed = current_filters != last_filter_state
                
                if filter_state_changed or row_count_changed:
                    # Get lightweight accurate count without UI manipulation
                    try:
                        updated_count = self.get_direct_visible_count()
                        if updated_count > 0:
                            current_row_count = updated_count
                    except:
                        pass
                        
                    # Filter change detected
                    from datetime import datetime
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # Find what changed
                    added_filters = []
                    removed_filters = []
                    filter_columns = []  # Track ALL columns that changed
                    
                    # Identify newly added or modified filters
                    if filter_state_changed:
                        for col, filter_info in current_filters.items():
                            # Get values or handle the new dictionary structure
                            if isinstance(filter_info, dict):
                                filter_values = filter_info.get("values", [])
                                column_index = filter_info.get("column_index", 0)
                            else:
                                # Handle older format for backward compatibility
                                filter_values = filter_info
                                column_index = 0
                                
                            if col not in last_filter_state:
                                added_filters.append({
                                    "column": col,
                                    "values": filter_values,
                                    "column_index": column_index
                                })
                                filter_columns.append(col)  # This is a newly added filter column
                            elif filter_values != (last_filter_state[col].get("values", []) 
                                                  if isinstance(last_filter_state[col], dict) 
                                                  else last_filter_state[col]):
                                added_filters.append({
                                    "column": col,
                                    "values": filter_values,
                                    "column_index": column_index
                                })
                                filter_columns.append(col)  # This filter column was modified
                                
                        # Identify removed filters
                        for col, filter_info in last_filter_state.items():
                            if col not in current_filters:
                                # Get values or handle the new dictionary structure
                                if isinstance(filter_info, dict):
                                    filter_values = filter_info.get("values", [])
                                    column_index = filter_info.get("column_index", 0)
                                else:
                                    # Handle older format for backward compatibility
                                    filter_values = filter_info
                                    column_index = 0
                                    
                                removed_filters.append({
                                    "column": col,
                                    "values": filter_values,
                                    "column_index": column_index
                                })
                                filter_columns.append(f"Remove {col}")  # This filter was removed
                        
                        # Update tracking of known filters
                        self.last_known_filters = {}
                        for col, values in current_filters.items():
                            self.last_known_filters[col] = values
                    
                    # If rows changed but no filter changes detected, try to detect which column changed
                    if row_count_changed and not filter_state_changed:
                        # First check if any autofilter is active
                        has_active_filter = False
                        try:
                            if hasattr(self.worksheet, "AutoFilter") and self.worksheet.AutoFilter:
                                for i in range(1, 100):  # Check reasonable number of columns
                                    try:
                                        filter_obj = self.worksheet.AutoFilter.Filters(i)
                                        if hasattr(filter_obj, "On") and filter_obj.On:
                                            has_active_filter = True
                                            
                                            # Try to get column name
                                            try:
                                                header_val = self.worksheet.Cells(1, i).Value
                                                if header_val:
                                                    col_name = str(header_val)
                                                    
                                                    # Check if this was a previously seen filter
                                                    if col_name in current_filters:
                                                        filter_columns.append(col_name)
                                                        
                                                        # Add to our added filters list
                                                        if col_name not in self.last_known_filters:
                                                            if isinstance(current_filters[col_name], dict):
                                                                filter_values = current_filters[col_name].get("values", [])
                                                                column_index = current_filters[col_name].get("column_index", i)
                                                            else:
                                                                filter_values = current_filters[col_name]
                                                                column_index = i
                                                                
                                                            added_filters.append({
                                                                "column": col_name,
                                                                "values": filter_values,
                                                                "column_index": column_index
                                                            })
                                                            self.last_known_filters[col_name] = current_filters[col_name]
                                                        # Or update if values changed
                                                        elif (isinstance(self.last_known_filters.get(col_name), dict) and 
                                                              isinstance(current_filters[col_name], dict) and
                                                              self.last_known_filters.get(col_name).get("values") != 
                                                              current_filters[col_name].get("values")) or (
                                                              not isinstance(self.last_known_filters.get(col_name), dict) and
                                                              self.last_known_filters.get(col_name) != current_filters[col_name]):
                                                              
                                                            if isinstance(current_filters[col_name], dict):
                                                                filter_values = current_filters[col_name].get("values", [])
                                                                column_index = current_filters[col_name].get("column_index", i)
                                                            else:
                                                                filter_values = current_filters[col_name]
                                                                column_index = i
                                                                
                                                            added_filters.append({
                                                                "column": col_name,
                                                                "values": filter_values,
                                                                "column_index": column_index
                                                            })
                                                            self.last_known_filters[col_name] = current_filters[col_name]
                                            except:
                                                pass
                                    except:
                                        continue
                        except:
                            pass
                            
                        # If we couldn't detect the specific column but filters are active
                        if len(filter_columns) == 0 and (has_active_filter or current_filters):
                            # Add all active filters to the columns list
                            if current_filters:
                                for col in current_filters.keys():
                                    filter_columns.append(col)
                    
                    # If we still don't have any filter columns but row count changed
                    if len(filter_columns) == 0 and row_count_changed:
                        filter_columns.append("Row Count Change")
                    
                    # Ensure we capture the specific column that changed
                    # Use a more aggressive approach to identify the changed column
                    
                    # Start with a specific default that's more meaningful
                    display_filter_column = "Filter Change"
                    
                    # Directly check active filter objects to get column names
                    if filter_state_changed:
                        try:
                            # Prioritize checking autofilter state directly to ensure accuracy
                            if hasattr(self.worksheet, "AutoFilter") and self.worksheet.AutoFilter:
                                for i in range(1, 200):  # Check a large number of columns
                                    try:
                                        filter_obj = self.worksheet.AutoFilter.Filters(i)
                                        if hasattr(filter_obj, "On") and filter_obj.On:
                                            # Get column name from header cache
                                            col_name = self.header_names.get(i)
                                            if not col_name:  # If not in cache, read directly
                                                cell_value = self.worksheet.Cells(1, i).Value
                                                col_name = str(cell_value) if cell_value else f"Column {i}"
                                                # Update cache for future use
                                                self.header_names[i] = col_name
                                                
                                            # Add to our changed columns list if not already there
                                            if col_name not in filter_columns:
                                                filter_columns.append(col_name)
                                    except:
                                        continue
                        except:
                            pass
                    
                    # For added or modified filters, use the most recently added/changed column
                    if added_filters:
                        display_filter_column = added_filters[0]["column"]
                    # For removed filters, indicate which was removed
                    elif removed_filters:
                        display_filter_column = f"Remove {removed_filters[0]['column']}"
                    # For row count changes without filter changes, use Row Count Change
                    elif row_count_changed and not filter_state_changed:
                        display_filter_column = "Row Count Change"
                    
                    # For a completely empty filter, show "No Filters"
                    if not current_filters and row_count_changed:
                        display_filter_column = "No Filters"
                        
                    # Debug log to help diagnose filter name issues
                    print(f"Changed columns: {filter_columns}")
                    print(f"Selected display column: {display_filter_column}")
                    
                    # Create simplified event with just most recent filter column
                    filter_event = {
                        "timestamp": timestamp,
                        "action": "filter_change",
                        "added_filters": added_filters,
                        "removed_filters": removed_filters,
                        "previous_row_count": last_row_count,
                        "current_row_count": current_row_count,
                        "total_rows": total_rows,
                        "filter_column": display_filter_column,       # Simplified to show only latest change
                        "filter_columns": filter_columns,             # Just tracking the columns that changed in this step
                        "active_filters": list(current_filters.keys()) # List of currently active filter columns
                    }
                    
                    print(f"\nFilter change detected at {timestamp}")
                    if added_filters:
                        print(f"Filters added/modified: {', '.join([f['column'] for f in added_filters])}")
                    if removed_filters:
                        print(f"Filters removed: {', '.join([f['column'] for f in removed_filters])}")
                    print(f"Rows: {last_row_count} -> {current_row_count} (of {total_rows})")
                    
                    # Only add events with significant changes
                    if added_filters or removed_filters or abs(last_row_count - current_row_count) > total_rows * 0.001:
                        self.filter_history.append(filter_event)
                        self.save_filter_history()
                    
                    last_filter_state = current_filters.copy()  # Make a copy to avoid reference issues
                    last_row_count = current_row_count
                
                # Reduced polling frequency to lower CPU usage
                time.sleep(0.25)  # Poll 4 times per second
                
        except Exception as e:
            print(f"Error in monitoring loop: {e}")
            self.save_filter_history()
            if status_queue:
                status_queue.put(f"Monitoring error: {str(e)}")
            return False
        finally:
            # Save one last time when stopping
            self.save_filter_history()
            # Uninitialize COM
            pythoncom.CoUninitialize()
            
        if status_queue:
            status_queue.put("Monitoring stopped")
            
        return True
    
    def get_direct_visible_count(self):
        """Get visible row count using primary key column for accurate matching with Excel"""
        import win32com.client
        
        try:
            # Get data range dimensions
            last_row = self.worksheet.UsedRange.Rows.Count
            
            # METHOD 1: Use SUBTOTAL on primary key column (most accurate match to Excel's count)
            try:
                # Use primary key column for counting
                primary_col = self.primary_key_column
                
                # Build range for primary key column (skip header row)
                pk_range = self.worksheet.Range(
                    self.worksheet.Cells(2, primary_col),  # Start from row 2 (after header)
                    self.worksheet.Cells(last_row, primary_col)  # End at last row
                )
                
                # Use SUBTOTAL(3, range) - COUNTA counts visible non-empty cells
                # This most closely matches what Excel shows as "x records found"
                count_formula = f"=SUBTOTAL(3,{pk_range.Address})"
                result = self.worksheet.Evaluate(count_formula)
                
                if result and isinstance(result, (int, float)) and result > 0:
                    print(f"Primary key count method: {int(result)} visible non-empty rows")
                    return int(result)  # This should match Excel's display exactly
            except Exception as e:
                print(f"Primary key method failed: {str(e)}")
                pass
            
            # METHOD 2: Try status bar reading (backup method)
            try:
                # Get the Excel status bar text
                status_text = self.excel.StatusBar
                
                if status_text and isinstance(status_text, str):
                    # Look for "x records found" or "x of y records found" patterns
                    import re
                    match = re.search(r'(\d+)(?:\s+of\s+\d+)?\s+record', status_text)
                    if match:
                        count = int(match.group(1))
                        print(f"Status bar method: {count} records")
                        return count
            except:
                pass
            
            # METHOD 3: Try SpecialCells on primary key column only
            try:
                # Use just the primary key column (skip header row)
                pk_range = self.worksheet.Range(
                    self.worksheet.Cells(2, self.primary_key_column),
                    self.worksheet.Cells(last_row, self.primary_key_column)
                )
                
                # Get visible cells in primary key column
                visibleCells = pk_range.SpecialCells(win32com.client.constants.xlCellTypeVisible)
                
                # Count non-empty cells only (to match Excel's behavior)
                non_empty_count = 0
                for cell in visibleCells:
                    if cell.Value is not None and cell.Value != "":
                        non_empty_count += 1
                
                print(f"Special cells method: {non_empty_count} visible non-empty rows")
                return non_empty_count
            except:
                pass
            
            # Last resort: total count if no filters seem active
            if not hasattr(self.worksheet, "FilterMode") or not self.worksheet.FilterMode:
                return self.get_total_row_count()
            else:
                # Default estimate - better than 0
                return max(1, int(self.get_total_row_count() * 0.5))
        except Exception as e:
            print(f"All count methods failed: {str(e)}")
            return 0
    
    def get_special_cells_count(self):
        """Get count using SpecialCells method - Excel's native visible cell detection"""
        import win32com.client
        
        try:
            # First try to get Excel's built-in visible row count
            # This is similar to what Excel shows in the "Records Found" status
            try:
                # Get the data range (excluding header row)
                dataRange = self.worksheet.Range(
                    self.worksheet.Cells(2, 1),  # Start from row 2 (after header)
                    self.worksheet.Cells(self.worksheet.UsedRange.Rows.Count, 1)
                )
                
                # Get only visible cells in the range
                visibleRange = dataRange.SpecialCells(win32com.client.constants.xlCellTypeVisible)
                
                # Count the row numbers
                rowNumbers = set()
                for cell in visibleRange.Cells:
                    rowNumbers.add(cell.Row)
                
                # Return the count of unique visible rows
                return len(rowNumbers)
            except:
                pass
                
            # Fallback to simpler SpecialCells method
            try:
                visible_cells = self.worksheet.UsedRange.SpecialCells(win32com.client.constants.xlCellTypeVisible)
                return max(0, visible_cells.Rows.Count - 1)  # Subtract header row
            except:
                return 0
        except:
            return 0
    
    def get_sampling_count(self):
        """Get count using row sampling method"""
        try:
            # Get total number of rows
            total_rows = self.get_total_row_count()
            
            # If there's no filter active, return total
            if not hasattr(self.worksheet, "FilterMode") or not self.worksheet.FilterMode:
                return total_rows
                
            # Sampling approach based on sheet size
            if total_rows > 10000:
                # For very large sheets, use sparse sampling
                sample_rate = max(1, total_rows // 500)  # Sample ~500 rows
            elif total_rows > 1000:
                # For medium sheets, sample more frequently
                sample_rate = max(1, total_rows // 200)  # Sample ~200 rows
            else:
                # For small sheets, check every 5th row
                sample_rate = 5
                
            # Count visible rows in sample
            visible_count = 0
            sample_count = 0
            
            for row in range(2, total_rows + 2, sample_rate):  # Start after header, +2 to include last section
                try:
                    if not self.worksheet.Rows(row).Hidden:
                        visible_count += 1
                    sample_count += 1
                except:
                    continue
                    
            # Calculate percentage and extrapolate
            if sample_count > 0:
                percentage = visible_count / sample_count
                estimated_count = int(percentage * total_rows)
                return estimated_count
            else:
                return 0
        except:
            return 0
            
    def start_monitoring(self, status_queue=None):
        """Start monitoring Excel filter operations"""
        # Reset stop event
        self.stop_event.clear()
        return self.background_monitoring(status_queue)
    
    def stop_monitoring(self):
        """Signal the monitoring thread to stop"""
        self.stop_event.set()
        self.save_filter_history()
        print("Stopping monitor...")
    
    def is_excel_alive(self):
        """Check if Excel is still responsive"""
        try:
            # Quick test to see if Excel is still alive
            _ = self.excel.Hwnd
            return True
        except:
            return False
            
    def get_current_filters(self):
        """Get the current filter state for all columns using cached header names"""
        import win32com.client
        
        filters = {}
        
        try:
            # Ensure we have header names cached
            if not self.header_names:
                self.read_all_headers()
            
            # FAST CHECK: Check if AutoFilter is available and active
            if (hasattr(self.worksheet, "AutoFilter") and 
                self.worksheet.AutoFilter is not None and 
                not isinstance(self.worksheet.AutoFilter, type(None))):
                
                # Get the filter range (just for boundary information)
                filter_range = self.worksheet.AutoFilter.Range
                
                # Get number of columns with a cap for performance
                try:
                    columns_count = min(filter_range.Columns.Count, 200)  # Allow up to 200 columns
                except:
                    columns_count = 50  # Fallback
                
                # Now quickly check which columns have active filters
                active_filters = {}
                for i in range(1, columns_count + 1):
                    try:
                        filter_obj = self.worksheet.AutoFilter.Filters(i)
                        if hasattr(filter_obj, "On") and filter_obj.On:
                            active_filters[i] = filter_obj
                    except:
                        continue
                
                # Now we only process active filters for better performance
                for i, filter_obj in active_filters.items():
                    try:
                        # Get column name from our cached headers - more reliable
                        column_name = self.header_names.get(i)
                        
                        # If header name wasn't cached for some reason, get it directly
                        if not column_name:
                            try:
                                cell_value = self.worksheet.Cells(1, i).Value
                                column_name = str(cell_value) if cell_value is not None else f"Column {i}"
                                # Add to cache for future use
                                self.header_names[i] = column_name
                            except:
                                column_name = f"Column {i}"
                        
                        # Get filter values
                        filter_values = []
                        
                        # Get filter criteria
                        if hasattr(filter_obj, "Criteria1") and filter_obj.Criteria1 is not None:
                            try:
                                filter_values.append(str(filter_obj.Criteria1))
                            except:
                                filter_values.append("Custom filter")
                                
                        if hasattr(filter_obj, "Criteria2") and filter_obj.Criteria2 is not None:
                            try:
                                filter_values.append(str(filter_obj.Criteria2))
                            except:
                                if not filter_values:
                                    filter_values.append("Custom filter")
                        
                        # If no values but filter is on, add a generic value
                        if not filter_values:
                            filter_values.append("Active")
                            
                        # Store in our filters dictionary with details needed for visualization
                        filters[column_name] = {
                            "values": filter_values,
                            "column_index": i,
                            "last_detected": time.time()
                        }
                    except Exception as e:
                        print(f"Error getting filter for column {i}: {str(e)}")
                        continue
                        
            # If AutoFilter approach failed but we have UsedRange, try looking for visible rows
            elif not filters and hasattr(self.worksheet, "UsedRange"):
                # This is a backup approach that doesn't get filter details
                # but can tell us something changed
                try:
                    used_range = self.worksheet.UsedRange
                    visible_count = self.get_visible_row_count()
                    total_count = self.get_total_row_count()
                    
                    # If visible count is different from total, assume some filtering is active
                    if visible_count < total_count:
                        filters["Unknown filter"] = {
                            "values": [f"Showing {visible_count} of {total_count} rows"],
                            "column_index": 0,
                            "last_detected": time.time()
                        }
                except:
                    pass
                
        except:
            # Don't flood the console with repeated errors
            pass
            
        return filters
        
    def get_visible_row_count(self):
        """Count visible rows (not filtered out) using a faster method"""
        import win32com.client
        
        try:
            # Try each method in order until one succeeds
            try:
                return self.get_direct_visible_count()
            except:
                pass
                
            try:
                return self.get_special_cells_count()
            except:
                pass
                
            try:
                return self.get_sampling_count()
            except:
                pass
                
            # Last resort: Return total rows if no filter seems active
            return self.get_total_row_count()
        except:
            # Return a default value
            return 0
            
    def get_primary_key_count(self):
        """Get count of non-empty cells in primary key column (excluding header)"""
        import win32com.client
        
        try:
            # Get data range dimensions
            last_row = self.worksheet.UsedRange.Rows.Count
            
            # Use primary key column for counting
            primary_col = self.primary_key_column
            
            # Build range for primary key column (skip header row)
            pk_range = self.worksheet.Range(
                self.worksheet.Cells(2, primary_col),  # Start from row 2 (after header)
                self.worksheet.Cells(last_row, primary_col)  # End at last row
            )
            
            # Use SUBTOTAL(3, range) - COUNTA counts non-empty cells
            # This most closely matches what Excel shows as "x records found"
            count_formula = f"=SUBTOTAL(3,{pk_range.Address})"
            result = self.worksheet.Evaluate(count_formula)
            
            if result and isinstance(result, (int, float)) and result > 0:
                return int(result)
            
            # Fallback: manually count non-empty cells
            non_empty_count = 0
            for row in range(2, last_row + 1):  # Skip header row
                try:
                    cell_value = self.worksheet.Cells(row, primary_col).Value
                    if cell_value is not None and cell_value != "":
                        non_empty_count += 1
                except:
                    continue
                    
            return non_empty_count
        except Exception as e:
            print(f"Error in get_primary_key_count: {e}")
            return 0
    
    def get_total_row_count(self):
        """Get total row count (excluding header)"""
        try:
            return max(0, self.worksheet.UsedRange.Rows.Count - 1)
        except:
            return 0
    
    def save_filter_history(self):
        """Save filter history to file"""
        try:
            with open(self.data_file, 'w') as f:
                json.dump(self.filter_history, f, indent=2)
        except Exception as e:
            print(f"Error saving filter history: {e}")
            
            # Try saving to a backup file
            try:
                backup_file = "filter_data_backup.json"
                with open(backup_file, 'w') as f:
                    json.dump(self.filter_history, f, indent=2)
                print(f"Saved to backup file: {backup_file}")
            except:
                pass

# Filter visualizer - generates sankey diagrams and tables
class FilterVisualizer:
    def __init__(self, data_file="filter_data.json"):
        self.data_file = data_file
        self.filter_history = []
        self.primary_key_column = 1  # Default to first column
        self.primary_key_name = ""   # Will be populated from monitor
        self.header_names = {}      # Will be populated from monitor
        
        # Create empty filter history file if it doesn't exist
        if not os.path.exists(self.data_file):
            with open(self.data_file, 'w') as f:
                json.dump([], f)
                
        self.load_filter_history()
        
    def load_filter_history(self):
        """Load filter history from file"""
        try:
            if os.path.exists(self.data_file) and os.path.getsize(self.data_file) > 0:
                with open(self.data_file, 'r') as f:
                    data = json.load(f)
                    
                if not isinstance(data, list):
                    print(f"ERROR: Invalid filter history data format. Expected list, got {type(data)}")
                    self.filter_history = []
                else:
                    self.filter_history = data
                    print(f"Loaded {len(self.filter_history)} filter events from {self.data_file}")
                    
                    # Debug: Print first few entries
                    if self.filter_history:
                        print(f"First event: {self.filter_history[0].get('action', 'unknown')} - {self.filter_history[0].get('filter_column', 'unknown')}")
            else:
                print(f"Filter history file is empty or does not exist: {self.data_file}")
                self.filter_history = []
        except Exception as e:
            print(f"Error loading filter history: {e}")
            import traceback
            traceback.print_exc()
            self.filter_history = []
            
            # Try backup file if it exists
            backup_file = "filter_data_backup.json"
            if os.path.exists(backup_file):
                try:
                    with open(backup_file, 'r') as f:
                        self.filter_history = json.load(f)
                    print(f"Loaded {len(self.filter_history)} filter events from backup file")
                except Exception as e:
                    print(f"Error loading backup filter history: {e}")
    
    def create_filter_table(self):
        """Create a table showing each filter step with column names and values"""
        import plotly.graph_objects as go
        import pandas as pd
        from plotly.graph_objects import Table
        
        if not self.filter_history:
            return None
            
        # Create dataframe for filter history
        data = []
        column_set = set()  # Track all unique columns that had filters
        
        # First pass to collect all column names
        for event in self.filter_history:
            if event.get("added_filters"):
                for filter_info in event["added_filters"]:
                    column_set.add(filter_info.get("column", "Unknown"))
            if event.get("removed_filters"):
                for filter_info in event["removed_filters"]:
                    column_set.add(filter_info.get("column", "Unknown"))
        
        # Sort columns for consistent display
        all_columns = sorted(list(column_set))
        
        # Track active filters at each step
        active_filters = {}  # column -> values
        
        # Track event index to name mapping for editing
        self.event_index_map = {}
        
        # Build data for each step
        for i, event in enumerate(self.filter_history):
            # Use the filter_column field if available
            filter_col = event.get("filter_column", None)
            if not filter_col and event.get("added_filters") and len(event["added_filters"]) > 0:
                filter_col = event["added_filters"][0]["column"]
            elif not filter_col and event.get("removed_filters") and len(event["removed_filters"]) > 0:
                filter_col = f"Remove {event['removed_filters'][0]['column']}"
                
            step_name = filter_col or f"Step {i+1}"
            
            # Store the mapping between step index and event index
            self.event_index_map[i] = {
                "event_index": self.filter_history.index(event),
                "filter_column": filter_col
            }
            
            row = {
                "Index": i,  # Store index for editing
                "Step": step_name,
                "Timestamp": event.get("timestamp", ""),
                "Rows": event.get("current_row_count", 0),
                "% of Total": f"{(event.get('current_row_count', 0) / event.get('total_rows', 1) * 100):.1f}%"
            }
            
            # Update active filters based on added/removed
            if event.get("added_filters"):
                for filter_info in event["added_filters"]:
                    col = filter_info.get("column", "Unknown")
                    # Extract values while handling different structures
                    if isinstance(filter_info.get("values"), dict):
                        vals = filter_info.get("values", {}).get("values", [])
                    else:
                        vals = filter_info.get("values", [])
                    active_filters[col] = vals
                    
            if event.get("removed_filters"):
                for filter_info in event["removed_filters"]:
                    col = filter_info.get("column", "Unknown")
                    if col in active_filters:
                        del active_filters[col]
            
            # Add all active filters to this row
            for col in all_columns:
                if col in active_filters:
                    vals = active_filters[col]
                    row[col] = ", ".join(vals) if isinstance(vals, list) else str(vals)
                else:
                    row[col] = ""
            
            data.append(row)
            
        # Create editable table with step names
        step_names = [row["Step"] for row in data]
        timestamps = [row["Timestamp"] for row in data]
        rows_counts = [row["Rows"] for row in data]
        percentages = [row["% of Total"] for row in data]
        
        # Generate background colors - make Step column editable (different color)
        editable_column_color = 'rgba(144, 238, 144, 0.3)'  # Light green for editable 
        normal_column_color = 'lavender'
        
        cell_colors = [
            [editable_column_color] * len(step_names),  # Step column is editable
            [normal_column_color] * len(timestamps),  
            [normal_column_color] * len(rows_counts),
            [normal_column_color] * len(percentages)
        ]
        
        # Add filter column values
        filter_values = []
        for col in all_columns:
            col_values = []
            for row in data:
                col_values.append(row.get(col, ""))
            filter_values.append(col_values)
        
        # Add colors for filter columns
        for _ in range(len(all_columns)):
            cell_colors.append([normal_column_color] * len(step_names))
        
        # Create table with hover text and special formatting for editable cells
        fig = go.Figure(data=[go.Table(
            header=dict(
                values=["Step<br><i>(click to edit)</i>", "Timestamp", "Rows", "% of Total"] + all_columns,
                fill_color='paleturquoise',
                align='left',
                font=dict(size=12)
            ),
            cells=dict(
                values=[step_names, timestamps, rows_counts, percentages] + filter_values,
                fill_color=cell_colors,
                align='left',
                font=dict(size=11),
                height=30
            ),
            columnwidth=[120, 120, 70, 70] + [80] * len(all_columns)
        )])
        
        # Add interactive capabilities with custom data
        # Store indexes in custom_data attribute for JavaScript interaction
        indexes = [row["Index"] for row in data]
        # Create custom HTML with JavaScript to handle cell editing
        # This HTML will embed the table visualization with added JavaScript
        # for handling clicks on the Step column cells
        
        fig.update_layout(
            title=dict(
                text="Filter Values by Column at Each Step<br><span style='font-size:12px;color:green'>Click on any Step name to edit it</span>",
                font=dict(size=14)
            ),
            margin=dict(l=10, r=10, t=60, b=10),
            height=400
        )
        
        # Add annotation to explain how to edit
        fig.add_annotation(
            x=0,
            y=1.05,
            xref="paper",
            yref="paper",
            text="Click on any Step name to edit the filter column name",
            showarrow=False,
            font=dict(size=10, color="green"),
            align="left"
        )
        
        return fig
            
    def create_sankey_diagram(self):
        """Create Sankey diagram from filter history showing each single change step"""
        import plotly.graph_objects as go
        
        if not self.filter_history:
            # Create empty figure with message
            fig = go.Figure()
            fig.add_annotation(
                text="No filter data detected yet.<br>Apply filters in Excel to see visualizations.",
                xref="paper", yref="paper",
                x=0.5, y=0.5,
                showarrow=False,
                font=dict(size=20)
            )
            return fig
        
        # Add title annotation with primary key information
        pk_title = f"<b>Filter Flow Visualization</b><br><i>Showing non-empty rows in {self.primary_key_name or f'Column {self.primary_key_column}'}</i>"
        
        # Start with all data
        total_rows = self.filter_history[0].get("total_rows", 1) 
        if total_rows <= 0:
            total_rows = 1  # Avoid division by zero
            
        # Create nodes - All Data is the first node
        nodes = ["All Data"]
        node_colors = ["rgba(31, 119, 180, 0.8)"]  # Blue for initial node
        
        # For keeping track of links
        links_source = []
        links_target = []
        links_value = []
        links_label = []
        links_color = []
        
        # Track active filters state for node naming
        active_filters = {}
        
        # Current row count
        current_row_count = total_rows
        
        for i, event in enumerate(self.filter_history):
            # Get the filter column name from the event - now it should always be just the latest changed column
            filter_col = event.get("filter_column", None)
            
            # Fall back to older method if needed
            if not filter_col:
                if event.get("added_filters") and len(event["added_filters"]) > 0:
                    filter_col = event["added_filters"][0]["column"]
                elif event.get("removed_filters") and len(event["removed_filters"]) > 0:
                    filter_col = f"Remove {event['removed_filters'][0]['column']}"
                else:
                    filter_col = f"Step {i+1}"
                
            # Update active filters state
            if event.get("added_filters"):
                for filter_info in event["added_filters"]:
                    col = filter_info.get("column", "Unknown")
                    vals = filter_info.get("values", [])
                    active_filters[col] = vals
                    
            if event.get("removed_filters"):
                for filter_info in event["removed_filters"]:
                    col = filter_info.get("column", "Unknown")
                    if col in active_filters:
                        del active_filters[col]
            
            # Create meaningful node name with filter info and counts
            row_count = event.get("current_row_count", 0)
            total = event.get("total_rows", 1)
            percentage = (row_count / total * 100) if total > 0 else 0
            step_name = f"{filter_col} ({row_count:,}, {percentage:.1f}%)"
                
            # Add node
            nodes.append(step_name)
            
            # Set node color based on number of remaining rows
            remaining_pct = event.get("current_row_count", 0) / total_rows if total_rows > 0 else 0
            if remaining_pct > 0.7:
                node_colors.append("rgba(50, 150, 50, 0.8)")  # Green for high
            elif remaining_pct > 0.3:
                node_colors.append("rgba(200, 150, 0, 0.8)")  # Yellow for medium
            else:
                node_colors.append("rgba(200, 50, 50, 0.8)")  # Red for low
            
            # Create link from previous state to this step
            links_source.append(i if i > 0 else 0)  # Previous node or All Data
            links_target.append(i+1)  # Current node
            
            # The value is how many rows flowed from previous step
            flow_value = max(1, event.get("current_row_count", 1))
            links_value.append(flow_value)
            
            # Create detailed label describing the filter changes
            filter_labels = []
            if event.get("added_filters"):
                for filter_info in event["added_filters"]:
                    col = filter_info.get("column", "Unknown")
                    vals = filter_info.get("values", [])
                    # Handle both list and dict formats
                    if isinstance(vals, dict) and "values" in vals:
                        vals = vals["values"]
                    val_str = ", ".join(vals) if isinstance(vals, list) else str(vals)
                    filter_labels.append(f"{col}: {val_str}")
            if event.get("removed_filters"):
                for filter_info in event["removed_filters"]:
                    col = filter_info.get("column", "Unknown")
                    filter_labels.append(f"Remove {col} filter")
            
            # Build comprehensive label
            if filter_labels:
                label = "<br>".join(filter_labels)
            else:
                label = f"Filter change {i+1}"
                
            # Add row counts to label
            prev_count = event.get("previous_row_count", 0)
            curr_count = event.get("current_row_count", 0)
            if prev_count > 0:
                change_pct = ((curr_count - prev_count) / prev_count * 100)
                change_str = f" ({change_pct:+.1f}%)" if change_pct != 0 else ""
            else:
                change_str = ""
            
            label += f"<br>Rows: {prev_count:,} → {curr_count:,}{change_str}"
            links_label.append(label)
            
            # Color based on percentage change from previous
            prev_count = event.get("previous_row_count", total_rows)
            if prev_count > 0:
                reduction = 1.0 - (flow_value / prev_count)
                # Red for high reduction, yellow for medium, green for low
                if reduction > 0.7:
                    links_color.append("rgba(255, 0, 0, 0.6)")  # Red
                elif reduction > 0.4:
                    links_color.append("rgba(255, 165, 0, 0.6)")  # Orange
                elif reduction > 0.1:
                    links_color.append("rgba(255, 255, 0, 0.6)")  # Yellow
                else:
                    links_color.append("rgba(0, 128, 0, 0.6)")  # Green
            else:
                links_color.append("rgba(100, 100, 100, 0.6)")  # Grey
            
            # Update current row count
            current_row_count = event.get("current_row_count", 0)
        
        # Create Sankey figure
        fig = go.Figure(data=[go.Sankey(
            node=dict(
                pad=20,
                thickness=25,
                line=dict(color="black", width=0.5),
                label=nodes,
                color=node_colors
            ),
            link=dict(
                source=links_source,
                target=links_target,
                value=links_value,
                label=links_label,
                color=links_color
            )
        )])
        
        fig.update_layout(
            title_text=pk_title,
            font=dict(size=14),
            height=600,
            margin=dict(l=25, r=25, t=80, b=25)  # More top margin for the title
        )
        
        return fig
    
    def save_full_report(self, filename):
        """Save a full HTML report with all visualizations"""
        import plotly.graph_objects as go
        import plotly.io as pio
        import pandas as pd
        
        try:
            # Create both visualizations
            sankey_fig = self.create_sankey_diagram()
            table_fig = self.create_filter_table() or go.Figure()
            
            # Write sankey figure to file
            pio.write_html(sankey_fig, file=filename, auto_open=False)
            print(f"Saved visualization to {filename}")
            return True
        except Exception as e:
            print(f"Error saving report: {e}")
            return False

# Thread for running Excel monitoring
class MonitorThread(QThread):
    """Thread to run Excel monitoring in background"""
    status_update = pyqtSignal(str)
    
    def __init__(self, monitor):
        super().__init__()
        self.monitor = monitor
        self.queue = None
        
    def run(self):
        import queue
        self.queue = queue.Queue()
        
        # Start monitoring in this thread
        try:
            self.monitor.start_monitoring(self.queue)
            
            # Process queue messages until monitor stops
            while True:
                try:
                    # Get message with timeout to allow checking if thread should stop
                    message = self.queue.get(timeout=0.5)
                    self.status_update.emit(message)
                except queue.Empty:
                    # Check if thread has been asked to stop
                    if self.isInterruptionRequested():
                        break
                except Exception as e:
                    self.status_update.emit(f"Error: {str(e)}")
        except Exception as e:
            self.status_update.emit(f"Monitor error: {str(e)}")
    
    def stop(self):
        """Stop the monitoring thread"""
        if self.monitor:
            self.monitor.stop_monitoring()
        self.requestInterruption()
        self.wait(2000)  # Wait up to 2 seconds

# Main application window
class CustomTableWidget(QWidget):
    """Custom table widget with row editing capability"""
    
    column_edited = pyqtSignal(int, str)  # Signal for when a column name is edited (row_index, new_name)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        # Import all required Qt classes at the beginning
        from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView, QLabel, QPushButton, QVBoxLayout, QHBoxLayout
        from PyQt5.QtGui import QColor, QFont, QBrush
        from PyQt5.QtCore import Qt
        
        # Store reference to required widget classes
        global QTableWidgetItem, QPushButton, QBrush, QColor, QFont, Qt
        QTableWidgetItem = QTableWidgetItem
        QPushButton = QPushButton
        QBrush = QBrush
        QColor = QColor
        QFont = QFont
        
        # Flag to prevent recursive updates
        self.updating_table = False
        
        # Setup layout
        self.layout = QVBoxLayout(self)
        
        # Create table widget
        self.table = QTableWidget()
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)  # Start with editing disabled (enabled after monitoring stops)
        self.table.cellChanged.connect(self.on_cell_changed)
        self.table.setColumnCount(5)  # Step, Timestamp, Rows, % of Total, Action
        self.table.setHorizontalHeaderLabels(["Step", "Timestamp", "Rows", "% of Total", "Action"])
        
        # Set column widths
        self.table.setColumnWidth(0, 150)  # Step column
        self.table.setColumnWidth(1, 180)  # Timestamp column
        self.table.setColumnWidth(2, 70)   # Rows column
        self.table.setColumnWidth(3, 80)   # % of Total column
        self.table.setColumnWidth(4, 70)   # Action column
        
        # Style the header
        header_font = QFont()
        header_font.setBold(True)
        self.table.horizontalHeader().setFont(header_font)
        self.table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #c0d9ea; }")
        
        # Other table settings
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget {
                alternate-background-color: #f0f0f0;
                background-color: white;
                selection-background-color: #a8d8ea;
                gridline-color: #d0d0d0;
            }
            QTableWidget::item:selected {
                color: black;
            }
        """)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        
        # Create action buttons
        from PyQt5.QtWidgets import QPushButton, QHBoxLayout
        button_layout = QHBoxLayout()
        
        self.refresh_btn = QPushButton("Refresh Table")
        self.refresh_btn.setStyleSheet("background-color: #e0e0e0; padding: 5px;")
        self.refresh_btn.clicked.connect(self.refresh_data)
        
        help_text = QLabel("Double-click on any Step name to edit")
        help_text.setStyleSheet("color: #008000; font-style: italic;")
        
        button_layout.addWidget(self.refresh_btn)
        button_layout.addWidget(help_text)
        button_layout.addStretch(1)
        
        # Add widgets to layout
        self.layout.addWidget(self.table)
        self.layout.addLayout(button_layout)
        
        # Data storage
        self.filter_history = []
        self.row_to_event_map = {}  # Maps table rows to filter history event indices
        
    def set_data(self, filter_history):
        """Set the filter history data and update the table"""
        self.filter_history = filter_history
        self.refresh_data()
        
    def refresh_data(self):
        """Refresh the table with current filter history data"""
        # Set flag to prevent recursive updates
        if self.updating_table:
            print("Table update already in progress, skipping to avoid recursion")
            return
            
        self.updating_table = True
        
        try:
            # Debug log to verify data
            print(f"Refreshing table with {len(self.filter_history) if self.filter_history else 0} events")
            
            # Temporarily disconnect cell changed signal to prevent recursive updates
            self.table.cellChanged.disconnect(self.on_cell_changed)
            
            self.table.setRowCount(0)  # Clear table
            self.row_to_event_map = {}
            
            if not self.filter_history:
                print("Warning: No filter history data to display")
                return
        except Exception as e:
            print(f"Error preparing table refresh: {e}")
            self.updating_table = False
            return
            
        # Populate table
        for i, event in enumerate(self.filter_history):
            row = self.table.rowCount()
            self.table.insertRow(row)
            
            # Map table row to event index
            self.row_to_event_map[row] = i
            
            # Get filter column
            filter_col = event.get("filter_column", None)
            if not filter_col and event.get("added_filters") and len(event["added_filters"]) > 0:
                filter_col = event["added_filters"][0]["column"]
            elif not filter_col and event.get("removed_filters") and len(event["removed_filters"]) > 0:
                filter_col = f"Remove {event['removed_filters'][0]['column']}"
                
            step_name = filter_col or f"Step {i+1}"
            
            # Add data to cells
            step_item = QTableWidgetItem(step_name)
            timestamp_item = QTableWidgetItem(event.get("timestamp", ""))
            rows_item = QTableWidgetItem(str(event.get("current_row_count", 0)))
            percent_item = QTableWidgetItem(f"{(event.get('current_row_count', 0) / event.get('total_rows', 1) * 100):.1f}%")
            
            # Make only the step name editable
            timestamp_item.setFlags(timestamp_item.flags() & ~Qt.ItemIsEditable)
            rows_item.setFlags(rows_item.flags() & ~Qt.ItemIsEditable)
            percent_item.setFlags(percent_item.flags() & ~Qt.ItemIsEditable)
            
            # Style the cells
            # Editable cell gets special styling
            step_item.setBackground(QBrush(QColor("#e6ffe6")))  # Light green background
            step_item.setToolTip("Double-click to edit this filter name")
            font = QFont()
            font.setBold(True)
            step_item.setFont(font)
            
            # Style the row based on event type
            action = event.get("action", "")
            row_color = QColor("white")
            
            if action == "initial_connection" or action == "initial_state":
                # Starting state - light blue
                row_color = QColor("#e6f7ff")
            elif event.get("current_row_count", 0) < event.get("previous_row_count", 0):
                # Row count decreased - light orange
                row_color = QColor("#fff2e6")
            
            # Apply colors to non-editable cells
            timestamp_item.setBackground(QBrush(row_color))
            rows_item.setBackground(QBrush(row_color))
            percent_item.setBackground(QBrush(row_color))
            
            # Center align numeric columns
            rows_item.setTextAlignment(Qt.AlignCenter)
            percent_item.setTextAlignment(Qt.AlignCenter)
            
            # Add items to table
            self.table.setItem(row, 0, step_item)
            self.table.setItem(row, 1, timestamp_item)
            self.table.setItem(row, 2, rows_item)
            self.table.setItem(row, 3, percent_item)
            
            # Add edit button in last column (using class-wide reference)
            edit_btn = QPushButton("Edit")
            edit_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    border: none;
                    border-radius: 3px;
                    padding: 5px;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
            """)
            edit_btn.clicked.connect(lambda checked, row=row: self.edit_row(row))
            self.table.setCellWidget(row, 4, edit_btn)
            
        # Reconnect cell changed signal and reset flag
        try:
            self.table.cellChanged.connect(self.on_cell_changed)
        except Exception as e:
            print(f"Error reconnecting signal: {e}")
        finally:
            self.updating_table = False
            
    def on_cell_changed(self, row, column):
        """Handle when a cell value is changed by the user"""
        # Skip if we're in the process of refreshing the table
        if self.updating_table:
            return
            
        if column == 0 and row in self.row_to_event_map:
            # Get the current value
            new_name = self.table.item(row, column).text()
            event_index = self.row_to_event_map[row]
            
            # Emit the signal for user edits
            print(f"User edited cell at row {row}, event index {event_index}: '{new_name}'")
            self.column_edited.emit(event_index, new_name)
            
    def edit_row(self, row):
        """Handle edit button click for a row"""
        if row in self.row_to_event_map:
            print(f"Editing row {row}, event index {self.row_to_event_map[row]}")
            self.table.editItem(self.table.item(row, 0))


class FilterTrailApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FilterTrail - Excel Filter Monitor")
        self.setGeometry(100, 100, 1000, 650)  # Smaller window size
        
        self.monitor = ExcelFilterMonitor()
        self.visualizer = FilterVisualizer()
        self.monitor_thread = None
        self.monitoring_active = False
        
        # Set primary key column (default: 1 = first column)
        self.primary_key_column = 1
        
        self.init_ui()
        
        # Set timer for auto-refresh if monitoring is active
        from PyQt5.QtCore import QTimer
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.auto_refresh_handler)
        self.auto_refresh_interval = 1000  # milliseconds - faster refresh (1 second)
        
    def init_ui(self):
        # Main widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # PRIMARY KEY SELECTION
        pk_area = QWidget()
        pk_layout = QHBoxLayout(pk_area)
        pk_layout.setContentsMargins(5, 5, 5, 5)
        
        # Primary key label
        pk_label = QLabel("Primary Key Column:")
        pk_label.setStyleSheet("font-weight: bold;")
        pk_layout.addWidget(pk_label)
        
        # Primary key selection (dropdown)
        from PyQt5.QtWidgets import QComboBox
        self.pk_selector = QComboBox()
        self.pk_selector.addItem("Column 1 (default)", 1)
        self.pk_selector.addItem("Column 2", 2)
        self.pk_selector.addItem("Column 3", 3)
        self.pk_selector.addItem("Column 4", 4)
        self.pk_selector.addItem("Column 5", 5)
        self.pk_selector.currentIndexChanged.connect(self.update_primary_key)
        pk_layout.addWidget(self.pk_selector)
        
        # Column name label (will be updated when connected)
        self.pk_name_label = QLabel("(First column)")
        pk_layout.addWidget(self.pk_name_label)
        
        # Add "Custom Name" button for primary key column
        from PyQt5.QtWidgets import QPushButton
        self.pk_name_btn = QPushButton("Rename")
        self.pk_name_btn.setMaximumWidth(80)
        self.pk_name_btn.clicked.connect(self.rename_primary_key)
        pk_layout.addWidget(self.pk_name_btn)
        
        # Add stretch to push elements to left
        pk_layout.addStretch(1)
        
        # Add to main layout
        main_layout.addWidget(pk_area)
        
        # COLUMN RENAMING SECTION
        rename_area = QWidget()
        rename_layout = QHBoxLayout(rename_area)
        rename_layout.setContentsMargins(5, 5, 5, 5)
        
        # Column rename label
        rename_label = QLabel("Rename Filter Column:")
        rename_label.setStyleSheet("font-weight: bold;")
        rename_layout.addWidget(rename_label)
        
        # Original column selection
        from PyQt5.QtWidgets import QComboBox, QLineEdit
        self.original_col_selector = QComboBox()
        self.original_col_selector.setMinimumWidth(150)
        rename_layout.addWidget(self.original_col_selector)
        
        # New name input
        rename_layout.addWidget(QLabel("→"))
        self.new_col_name = QLineEdit()
        self.new_col_name.setPlaceholderText("New column name")
        self.new_col_name.setMinimumWidth(150)
        rename_layout.addWidget(self.new_col_name)
        
        # Apply button
        self.rename_col_btn = QPushButton("Apply")
        self.rename_col_btn.setMaximumWidth(80)
        self.rename_col_btn.clicked.connect(self.rename_filter_column)
        rename_layout.addWidget(self.rename_col_btn)
        
        # Add stretch to push elements to left
        rename_layout.addStretch(1)
        
        # Add to main layout
        main_layout.addWidget(rename_area)
        
        # Load existing data for visualizer
        self.visualizer.load_filter_history()
        
        # Update filter column dropdown with current data
        self.update_filter_column_selector()
        
        # Status area
        status_area = QWidget()
        status_layout = QHBoxLayout(status_area)
        status_layout.setContentsMargins(5, 5, 5, 5)  # Reduce margins
        
        # Status label
        self.status_label = QLabel("Ready to connect to Excel")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-weight: bold; color: #333;")
        status_layout.addWidget(self.status_label)
        
        # Progress bar for visual indication of activity
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumWidth(200)  # Limit width
        status_layout.addWidget(self.progress_bar)
        
        main_layout.addWidget(status_area)
        
        # Create splitter for resizable areas
        splitter = QSplitter(Qt.Vertical)
        splitter.setHandleWidth(10)
        main_layout.addWidget(splitter, 1)
        
        # Web view for Sankey diagram
        sankey_widget = QWidget()
        sankey_layout = QVBoxLayout(sankey_widget)
        sankey_layout.setContentsMargins(5, 5, 5, 5)  # Reduce margins
        sankey_label = QLabel("Excel Filter Flow")
        sankey_label.setAlignment(Qt.AlignCenter)
        sankey_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        sankey_layout.addWidget(sankey_label)
        
        self.sankey_view = QWebEngineView()
        self.sankey_view.setMinimumHeight(300)
        sankey_layout.addWidget(self.sankey_view)
        
        # CUSTOM TABLE for filter steps with editing capability
        table_widget = QWidget()
        table_layout = QVBoxLayout(table_widget)
        table_layout.setContentsMargins(5, 5, 5, 5)  # Reduce margins
        table_label = QLabel("Filter Steps (Click to Edit Column Names)")
        table_label.setAlignment(Qt.AlignCenter)
        table_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        table_layout.addWidget(table_label)
        
        # Create custom table widget for direct editing
        self.filter_table = CustomTableWidget()
        self.filter_table.column_edited.connect(self.on_filter_column_edited)
        table_layout.addWidget(self.filter_table)
        
        # Keep the original web view for compatibility but hide it
        self.table_view = QWebEngineView()
        self.table_view.setVisible(False)
        self.table_view.setMaximumHeight(1)
        
        # Add both visualization areas to splitter
        splitter.addWidget(sankey_widget)
        splitter.addWidget(table_widget)
        
        # Set initial sizes (70% for Sankey, 30% for table)
        splitter.setSizes([700, 300])
        
        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setContentsMargins(5, 5, 5, 5)  # Reduce margins
        
        # Initialize table with existing filter history data
        if self.visualizer.filter_history:
            print(f"Initial table update with {len(self.visualizer.filter_history)} records")
            self.filter_table.set_data(self.visualizer.filter_history)
        
        self.start_btn = QPushButton("Start Monitoring")
        self.start_btn.clicked.connect(self.start_monitoring)
        self.start_btn.setMinimumWidth(120)
        btn_layout.addWidget(self.start_btn)
        
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.clicked.connect(self.stop_monitoring)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setMinimumWidth(80)
        btn_layout.addWidget(self.stop_btn)
        
        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.clicked.connect(self.refresh_visualization)
        self.refresh_btn.setMinimumWidth(80)
        btn_layout.addWidget(self.refresh_btn)
        
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save_visualization)
        self.save_btn.setMinimumWidth(80)
        btn_layout.addWidget(self.save_btn)
        
        self.reset_btn = QPushButton("Reset Data")
        self.reset_btn.clicked.connect(self.reset_data)
        self.reset_btn.setMinimumWidth(80)
        btn_layout.addWidget(self.reset_btn)
        
        main_layout.addLayout(btn_layout)
        
        # Use a single-shot timer to load initial data after UI is fully set up
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(100, self.delayed_initial_load)
        
    def delayed_initial_load(self):
        """Load initial data after UI is fully setup to avoid recursive issues"""
        # Load visualization data
        if hasattr(self.visualizer, 'filter_history'):
            self.visualizer.load_filter_history()
            
        # Populate table with initial data if available
        if hasattr(self.visualizer, 'filter_history') and self.visualizer.filter_history:
            if hasattr(self, 'filter_table'):
                self.filter_table.set_data(self.visualizer.filter_history)
                
        # Load sankey diagram
        self.refresh_visualization()
            
    def auto_refresh_handler(self):
        """Special handler for auto-refresh timer to ensure both Sankey and table update"""
        # Only update if we're actively monitoring
        if not self.monitoring_active:
            return
            
        # Check if filter data has changed before triggering refresh
        current_data_size = 0
        if hasattr(self.visualizer, "filter_history") and self.visualizer.filter_history:
            current_data_size = len(self.visualizer.filter_history)
        elif hasattr(self.monitor, "filter_history") and self.monitor.filter_history:
            current_data_size = len(self.monitor.filter_history)
            
        # Only update if we have data and if we're actively monitoring
        if current_data_size > 0:
            self.refresh_visualization()
        
    def update_primary_key(self):
        """Update the primary key column when selection changes"""
        # Get the selected primary key column index
        index = self.pk_selector.currentIndex()
        self.primary_key_column = self.pk_selector.itemData(index)
        
        # Update monitor's primary key column
        self.monitor.primary_key_column = self.primary_key_column
        
        # Update label if we're connected to Excel
        if self.monitor.worksheet:
            # Try to get column name from cached headers
            column_name = self.monitor.header_names.get(self.primary_key_column, f"Column {self.primary_key_column}")
            self.pk_name_label.setText(f"({column_name})")
            self.monitor.primary_key_name = column_name
            
            # Show status message
            self.status_label.setText(f"Primary key set to column {self.primary_key_column}: {column_name}")
            
    def start_monitoring(self):
        """Start Excel monitoring in a separate thread"""
        if self.monitoring_active:
            return
        
        self.progress_bar.setVisible(True)    
        self.status_label.setText("Connecting to Excel...")
        
        # Set primary key column from UI
        self.monitor.primary_key_column = self.primary_key_column
        
        # Disable table editing during monitoring
        if hasattr(self, 'filter_table'):
            # Disable editing using the table's own constant
            self.filter_table.table.setEditTriggers(self.filter_table.table.NoEditTriggers)
            
            # Change table style to indicate non-editable mode
            self.filter_table.table.setStyleSheet("""
                QTableWidget {
                    alternate-background-color: #f0f0f0;
                    background-color: white;
                    selection-background-color: #a8d8ea;
                    gridline-color: #d0d0d0;
                    border: 1px solid #cccccc;  /* Regular border during monitoring */
                }
                QTableWidget::item:selected {
                    color: black;
                }
                QHeaderView::section {
                    background-color: #c0d9ea;
                }
            """)
            
            # Update the refresh button to standard mode
            self.filter_table.refresh_btn.setText("Refresh")
            self.filter_table.refresh_btn.setStyleSheet("background-color: #e0e0e0; padding: 5px;")
        
        # Create new thread for monitoring
        self.monitor_thread = MonitorThread(self.monitor)
        self.monitor_thread.status_update.connect(self.update_status)
        self.monitor_thread.finished.connect(self.monitoring_finished)
        
        # Start the thread
        self.monitor_thread.start()
        
        # Update UI
        self.monitoring_active = True
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.pk_selector.setEnabled(False)  # Disable primary key selection while monitoring
        
        # Start auto-refresh timer
        self.refresh_timer.start(self.auto_refresh_interval)
        
    def stop_monitoring(self):
        """Stop the monitoring thread"""
        if not self.monitoring_active:
            return
            
        self.status_label.setText("Stopping monitor...")
        
        # Stop the thread
        if self.monitor_thread:
            self.monitor_thread.stop()
        
        # Stop auto-refresh timer
        self.refresh_timer.stop()
        
    def monitoring_finished(self):
        """Called when monitoring thread finishes"""
        self.monitoring_active = False
        self.progress_bar.setVisible(False)
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.pk_selector.setEnabled(True)  # Re-enable primary key selection
        
        # Enable the edit buttons in the filter table (only when monitoring is stopped)
        if hasattr(self, 'filter_table'):
            # Show a message in the status bar indicating editing is now available
            self.status_label.setText("Monitoring stopped - You can now edit step names in the table")
            
            # Make the table's cells editable using the table's own constants
            table = self.filter_table.table
            table.setEditTriggers(
                table.DoubleClicked | 
                table.SelectedClicked | 
                table.EditKeyPressed
            )
            
            # Change table style to indicate it's now editable
            self.filter_table.table.setStyleSheet("""
                QTableWidget {
                    alternate-background-color: #f0f0f0;
                    background-color: white;
                    selection-background-color: #a8d8ea;
                    gridline-color: #d0d0d0;
                    border: 2px solid #4CAF50;  /* Green border to indicate editable */
                }
                QTableWidget::item:selected {
                    color: black;
                }
                QHeaderView::section {
                    background-color: #c0d9ea;
                }
            """)
            
            # Update the refresh button to indicate editing mode
            self.filter_table.refresh_btn.setText("Refresh & Save Changes")
            self.filter_table.refresh_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    padding: 5px;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
            """)
        else:
            self.status_label.setText("Monitoring stopped")
        
        # Update primary key label with column names read from Excel
        if self.monitor and self.monitor.header_names:
            column_name = self.monitor.header_names.get(self.primary_key_column, f"Column {self.primary_key_column}")
            self.pk_name_label.setText(f"({column_name})")
            
            # Also update dropdown with actual column names
            self.update_column_dropdown()
        
        # Refresh visualization one more time
        self.refresh_visualization()
        
    def update_column_dropdown(self):
        """Update the primary key dropdown with actual column names from Excel"""
        if not self.monitor or not self.monitor.header_names:
            return
            
        # Save current selection
        current_pk = self.primary_key_column
        
        # Clear dropdown
        self.pk_selector.clear()
        
        # Add first 10 columns with their actual names
        for i in range(1, min(11, max(self.monitor.header_names.keys()) + 1)):
            col_name = self.monitor.header_names.get(i, f"Column {i}")
            display_text = f"Column {i}: {col_name}"
            self.pk_selector.addItem(display_text, i)
            
        # Restore selection
        index = self.pk_selector.findData(current_pk)
        if index >= 0:
            self.pk_selector.setCurrentIndex(index)
            
        # Also update the filter column selector
        self.update_filter_column_selector()
            
    def on_filter_column_edited(self, event_index, new_name):
        """Handle when a filter column name is edited in the table"""
        if not self.monitor:
            return
        
        try:
            # Get the event from history (use visualizer's history if available, otherwise monitor's)
            event = None
            original_name = ""
            
            # Try to get event from visualizer first
            if hasattr(self.visualizer, "filter_history") and self.visualizer.filter_history:
                if 0 <= event_index < len(self.visualizer.filter_history):
                    event = self.visualizer.filter_history[event_index]
                    original_name = event.get("filter_column", "")
            
            # If not found, try monitor's history
            if event is None and hasattr(self.monitor, "filter_history") and self.monitor.filter_history:
                if 0 <= event_index < len(self.monitor.filter_history):
                    event = self.monitor.filter_history[event_index]
                    original_name = event.get("filter_column", "")
            
            # If we couldn't find the event, exit
            if event is None:
                print(f"Could not find event at index {event_index} in history")
                return
                
            if not original_name or not new_name:
                return
            
            # Skip special entries
            if original_name in ["All Data", "Initial State"]:
                self.status_label.setText(f"Cannot rename special filter step: {original_name}")
                # Refresh to restore original name
                self.refresh_visualization()
                return
                
            # Update filter_column in this specific event
            event["filter_column"] = new_name
            
            # Update in other fields if present
            if "filter_columns" in event:
                for i, col in enumerate(event["filter_columns"]):
                    if col == original_name:
                        event["filter_columns"][i] = new_name
            
            # Also update in added_filters and removed_filters if present
            for filter_list in ["added_filters", "removed_filters"]:
                if filter_list in event:
                    for filter_info in event[filter_list]:
                        if filter_info.get("column") == original_name:
                            filter_info["column"] = new_name
            
            # Update header names if this matches a header
            if hasattr(self.monitor, "header_names"):
                for col_idx, col_name in self.monitor.header_names.items():
                    if col_name == original_name:
                        self.monitor.header_names[col_idx] = new_name
            
            # IMPORTANT: Update both visualizer and monitor filter_history to keep in sync
            # First save changes to monitor's history
            if hasattr(self.monitor, "filter_history"):
                # If edited in visualizer, copy changes to monitor
                if event in self.visualizer.filter_history and event not in self.monitor.filter_history:
                    # Find the corresponding event in monitor's history or copy the whole visualizer history
                    for i, mon_event in enumerate(self.monitor.filter_history):
                        if mon_event.get("timestamp") == event.get("timestamp"):
                            self.monitor.filter_history[i] = event.copy()
                            break
                
                # Save monitor's history to file
                self.monitor.save_filter_history()
                
            # Also update visualizer's history if it's different
            if hasattr(self.visualizer, "filter_history"):
                # If edited in monitor, copy changes to visualizer
                if event in self.monitor.filter_history and event not in self.visualizer.filter_history:
                    # Find the corresponding event in visualizer's history or copy the whole monitor history
                    for i, vis_event in enumerate(self.visualizer.filter_history):
                        if vis_event.get("timestamp") == event.get("timestamp"):
                            self.visualizer.filter_history[i] = event.copy()
                            break
                
                # Save to file if visualizer has a save method
                if hasattr(self.visualizer, "save_filter_history"):
                    self.visualizer.save_filter_history()
            
            # IMPORTANT: Refresh Sankey diagram to show updated names
            self.refresh_sankey_diagram()
            
            # Show confirmation
            self.status_label.setText(f"Renamed '{original_name}' to '{new_name}' in step {event_index+1}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"Error updating filter column name: {e}")
            self.status_label.setText(f"Error updating filter column name: {str(e)}")
    
    def refresh_sankey_diagram(self):
        """Refresh only the Sankey diagram without updating table"""
        try:
            import plotly.io as pio
            
            # Use simple consistent filename
            temp_sankey_html = "temp_sankey.html"
            
            # Make sure visualizer has the latest data
            self.visualizer.load_filter_history()
            
            # Pass primary key information to visualizer
            self.visualizer.primary_key_column = self.primary_key_column
            self.visualizer.primary_key_name = self.monitor.primary_key_name if self.monitor else ""
            
            # Create and save Sankey diagram
            sankey_fig = self.visualizer.create_sankey_diagram()
            if sankey_fig:
                # Write to simple fixed filename
                pio.write_html(sankey_fig, file=temp_sankey_html, auto_open=False, include_plotlyjs='cdn')
                # Ensure file exists and has correct permissions
                if os.path.exists(temp_sankey_html):
                    try:
                        # Set read permissions for everyone
                        os.chmod(temp_sankey_html, 0o644)
                    except:
                        pass
                # Load with direct URL 
                self.sankey_view.load(QUrl.fromLocalFile(os.path.abspath(temp_sankey_html)))
                
                # Force reload to ensure fresh content
                self.sankey_view.reload()
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.status_label.setText(f"Error refreshing Sankey diagram: {str(e)}")
            
    def update_filter_column_selector(self):
        """Update the filter column selector with all column names from history"""
        if not self.monitor:
            return
            
        # Clear dropdown
        self.original_col_selector.clear()
        
        # Collect unique column names from filter history
        column_set = set()
        
        # Add column names from header cache
        if hasattr(self.monitor, "header_names") and self.monitor.header_names:
            for col_idx, col_name in self.monitor.header_names.items():
                column_set.add(col_name)
                
        # Add column names from filter history
        if hasattr(self.monitor, "filter_history") and self.monitor.filter_history:
            for event in self.monitor.filter_history:
                if "filter_column" in event and event["filter_column"]:
                    column_set.add(event["filter_column"])
                if "filter_columns" in event and event["filter_columns"]:
                    for col in event["filter_columns"]:
                        column_set.add(col)
                        
        # Also check the visualizer's filter history
        if hasattr(self.visualizer, "filter_history") and self.visualizer.filter_history:
            for event in self.visualizer.filter_history:
                if "filter_column" in event and event["filter_column"]:
                    column_set.add(event["filter_column"])
                if "filter_columns" in event and event["filter_columns"]:
                    for col in event["filter_columns"]:
                        column_set.add(col)
        
        # Add all column names to the dropdown (sorted alphabetically)
        sorted_columns = sorted(list(column_set))
        for col_name in sorted_columns:
            # Skip generic names and special entries
            if col_name not in ["All Data", "No Filters", "Initial State", "Row Count Change"]:
                self.original_col_selector.addItem(col_name)
                
        print(f"Updated filter column selector with {len(sorted_columns)} columns")
        
    def rename_filter_column(self):
        """Rename a filter column throughout all filter history"""
        if not self.monitor:
            return
            
        # Get selected original column name
        original_name = self.original_col_selector.currentText()
        
        # Get new name
        new_name = self.new_col_name.text()
        
        if not original_name or not new_name:
            self.status_label.setText("Please select a column and enter a new name")
            return
            
        # Update header names if this is an actual column
        modified_count = 0
        
        if hasattr(self.monitor, "header_names"):
            for col_idx, col_name in self.monitor.header_names.items():
                if col_name == original_name:
                    self.monitor.header_names[col_idx] = new_name
                    modified_count += 1
        
        # Update filter history in monitor
        if hasattr(self.monitor, "filter_history") and self.monitor.filter_history:
            for event in self.monitor.filter_history:
                # Update main filter_column field
                if event.get("filter_column") == original_name:
                    event["filter_column"] = new_name
                    modified_count += 1
                    
                # Update in filter_columns list
                if "filter_columns" in event:
                    for i, col in enumerate(event["filter_columns"]):
                        if col == original_name:
                            event["filter_columns"][i] = new_name
                            modified_count += 1
                            
                # Update in added_filters and removed_filters
                for filter_list in ["added_filters", "removed_filters"]:
                    if filter_list in event:
                        for filter_info in event[filter_list]:
                            if filter_info.get("column") == original_name:
                                filter_info["column"] = new_name
                                modified_count += 1
        
        # Also update filter history in visualizer to ensure consistency
        if self.visualizer and hasattr(self.visualizer, "filter_history") and self.visualizer.filter_history:
            for event in self.visualizer.filter_history:
                # Update main filter_column field
                if event.get("filter_column") == original_name:
                    event["filter_column"] = new_name
                    modified_count += 1
                    
                # Update in filter_columns list
                if "filter_columns" in event:
                    for i, col in enumerate(event["filter_columns"]):
                        if col == original_name:
                            event["filter_columns"][i] = new_name
                            modified_count += 1
                            
                # Update in added_filters and removed_filters
                for filter_list in ["added_filters", "removed_filters"]:
                    if filter_list in event:
                        for filter_info in event[filter_list]:
                            if filter_info.get("column") == original_name:
                                filter_info["column"] = new_name
                                modified_count += 1
        
        # Save changes
        self.monitor.save_filter_history()
        
        # Update UI
        self.update_filter_column_selector()
        self.new_col_name.clear()
        
        # Refresh both table and Sankey visualization
        self.refresh_visualization()
        
        # Show confirmation
        self.status_label.setText(f"Renamed '{original_name}' to '{new_name}' ({modified_count} occurrences updated)")
    
    def rename_primary_key(self):
        """Rename the primary key column"""
        if not self.monitor:
            return
            
        from PyQt5.QtWidgets import QInputDialog
        
        # Get current primary key name
        current_name = self.monitor.primary_key_name or f"Column {self.primary_key_column}"
        
        # Show input dialog
        new_name, ok = QInputDialog.getText(
            self, 
            "Rename Primary Key Column",
            "Enter new name for the primary key column:",
            text=current_name
        )
        
        if ok and new_name:
            # Update the name
            self.monitor.primary_key_name = new_name
            self.pk_name_label.setText(f"({new_name})")
            
            # Update header names cache
            self.monitor.header_names[self.primary_key_column] = new_name
            
            # Refresh visualization to show new name
            self.refresh_visualization()
            
            # Show confirmation
            self.status_label.setText(f"Primary key column renamed to: {new_name}")
            
    def rename_filter_column(self):
        """Rename a filter column"""
        if not self.monitor:
            return
            
        # Get selected original column name
        original_name = self.original_col_selector.currentText()
        
        # Get new name
        new_name = self.new_col_name.text()
        
        if not original_name or not new_name:
            self.status_label.setText("Please select a column and enter a new name")
            return
            
        # Update header names if this is an actual column
        for col_idx, col_name in self.monitor.header_names.items():
            if col_name == original_name:
                self.monitor.header_names[col_idx] = new_name
        
        # Update filter history
        modified_count = 0
        if self.monitor.filter_history:
            for event in self.monitor.filter_history:
                # Update main filter_column field
                if event.get("filter_column") == original_name:
                    event["filter_column"] = new_name
                    modified_count += 1
                    
                # Update in filter_columns list
                if "filter_columns" in event:
                    for i, col in enumerate(event["filter_columns"]):
                        if col == original_name:
                            event["filter_columns"][i] = new_name
                            modified_count += 1
                            
                # Update in added_filters and removed_filters
                for filter_list in ["added_filters", "removed_filters"]:
                    if filter_list in event:
                        for filter_info in event[filter_list]:
                            if filter_info.get("column") == original_name:
                                filter_info["column"] = new_name
                                modified_count += 1
        
        # Save changes
        self.monitor.save_filter_history()
        
        # Update UI
        self.update_filter_column_selector()
        self.new_col_name.clear()
        
        # Refresh visualization
        self.refresh_visualization()
        
        # Show confirmation
        self.status_label.setText(f"Renamed '{original_name}' to '{new_name}' ({modified_count} occurrences updated)")
        
    def monitoring_finished(self):
        """Called when monitoring thread finishes"""
        self.monitoring_active = False
        self.progress_bar.setVisible(False)
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.pk_selector.setEnabled(True)  # Re-enable primary key selection
        
        # Enable the edit buttons in the filter table (only when monitoring is stopped)
        if hasattr(self, 'filter_table'):
            # Show a message in the status bar indicating editing is now available
            self.status_label.setText("Monitoring stopped - You can now edit step names in the table")
            
            # Make the table's cells editable using the table's own constants
            table = self.filter_table.table
            table.setEditTriggers(
                table.DoubleClicked | 
                table.SelectedClicked | 
                table.EditKeyPressed
            )
            
            # Change table style to indicate it's now editable
            self.filter_table.table.setStyleSheet("""
                QTableWidget {
                    alternate-background-color: #f0f0f0;
                    background-color: white;
                    selection-background-color: #a8d8ea;
                    gridline-color: #d0d0d0;
                    border: 2px solid #4CAF50;  /* Green border to indicate editable */
                }
                QTableWidget::item:selected {
                    color: black;
                }
                QHeaderView::section {
                    background-color: #c0d9ea;
                }
            """)
            
            # Update the refresh button to indicate editing mode
            self.filter_table.refresh_btn.setText("Refresh & Save Changes")
            self.filter_table.refresh_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    padding: 5px;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
            """)
        else:
            self.status_label.setText("Monitoring stopped")
        
        # Update primary key label with column names read from Excel
        if self.monitor and self.monitor.header_names:
            column_name = self.monitor.header_names.get(self.primary_key_column, f"Column {self.primary_key_column}")
            self.pk_name_label.setText(f"({column_name})")
            
            # Also update dropdown with actual column names
            self.update_column_dropdown()
        
        # Refresh visualization one more time
        self.refresh_visualization()
        
    def update_status(self, message):
        """Update status from monitoring thread"""
        self.status_label.setText(message)
        
    def reset_data(self):
        """Reset filter history data"""
        reply = QMessageBox.question(
            self, "Reset Data",
            "This will clear all filter history data. Are you sure?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # Stop monitoring if active
            if self.monitoring_active:
                self.stop_monitoring()
                
            # Clear data files
            try:
                data_file = "filter_data.json"
                backup_file = "filter_data_backup.json"
                
                # Write empty JSON array to files
                with open(data_file, 'w') as f:
                    f.write("[]")
                
                if os.path.exists(backup_file):
                    with open(backup_file, 'w') as f:
                        f.write("[]")
                
                # Refresh visualization
                self.refresh_visualization()
                
                QMessageBox.information(self, "Reset Complete", "Filter history has been reset.")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not reset data: {str(e)}")
        
    def refresh_visualization(self):
        """Refresh the Sankey diagram and table visualizations"""
        try:
            import plotly.io as pio
            
            # Use simple consistent filenames to avoid access problems
            temp_sankey_html = "temp_sankey.html"
            
            # Make sure we have the most up-to-date data
            self.visualizer.load_filter_history()
            
            # Make sure our monitor has the same data for consistency
            if hasattr(self.monitor, "filter_history") and self.visualizer.filter_history:
                self.monitor.filter_history = self.visualizer.filter_history.copy()
            
            # Pass primary key information to visualizer
            self.visualizer.primary_key_column = self.primary_key_column
            self.visualizer.primary_key_name = self.monitor.primary_key_name if self.monitor else ""
            
            # Pass the full header names dictionary to visualizer for consistent naming
            if self.monitor and hasattr(self.monitor, "header_names"):
                self.visualizer.header_names = self.monitor.header_names
            
            # Create and save Sankey diagram
            sankey_fig = self.visualizer.create_sankey_diagram()
            if sankey_fig:
                # Write to simple fixed filename
                pio.write_html(sankey_fig, file=temp_sankey_html, auto_open=False, include_plotlyjs='cdn')
                # Ensure file exists and has correct permissions
                if os.path.exists(temp_sankey_html):
                    try:
                        # Set read permissions for everyone
                        os.chmod(temp_sankey_html, 0o644)
                    except:
                        pass
                # Load with direct URL 
                self.sankey_view.load(QUrl.fromLocalFile(os.path.abspath(temp_sankey_html)))
                
                # Force reload to ensure fresh content
                self.sankey_view.reload()
            
            # IMPORTANT: Always update the table with the most current data
            # First check if we have filter history data available
            filter_history = None
            
            if hasattr(self.visualizer, "filter_history") and self.visualizer.filter_history:
                filter_history = self.visualizer.filter_history
                print(f"Updating table with {len(filter_history)} visualizer filter history records")
            elif hasattr(self.monitor, "filter_history") and self.monitor.filter_history:
                filter_history = self.monitor.filter_history
                print(f"Updating table with {len(filter_history)} monitor filter history records")
            
            # Update the table if we have data
            if filter_history:
                if hasattr(self, 'filter_table'):
                    # Force update the table with the current data
                    self.filter_table.set_data(filter_history)
                    print(f"Table updated with {len(filter_history)} records")
                else:
                    print("Warning: filter_table attribute not found")
            else:
                print("No filter history data available to update table")
                
            # Apply styling based on monitoring state
            if hasattr(self, 'filter_table'):
                if self.monitoring_active:
                    # During monitoring, disable editing
                    self.filter_table.table.setEditTriggers(self.filter_table.table.NoEditTriggers)
                    self.filter_table.table.setStyleSheet("""
                        QTableWidget {
                            alternate-background-color: #f0f0f0;
                            background-color: white;
                            selection-background-color: #a8d8ea;
                            gridline-color: #d0d0d0;
                            border: 1px solid #cccccc;
                        }
                        QTableWidget::item:selected {
                            color: black;
                        }
                    """)
                else:
                    # After monitoring stops, enable editing with visual cues
                    table = self.filter_table.table
                    self.filter_table.table.setEditTriggers(
                        table.DoubleClicked | table.SelectedClicked | table.EditKeyPressed
                    )
                    self.filter_table.table.setStyleSheet("""
                        QTableWidget {
                            alternate-background-color: #f0f0f0;
                            background-color: white;
                            selection-background-color: #a8d8ea;
                            gridline-color: #d0d0d0;
                            border: 2px solid #4CAF50;
                        }
                        QTableWidget::item:selected {
                            color: black;
                        }
                    """)
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.status_label.setText(f"Error refreshing visualization: {str(e)}")
            
    def save_visualization(self):
        """Save the current visualization to a file"""
        try:
            filename, _ = QFileDialog.getSaveFileName(
                self, "Save Visualization", "", "HTML Files (*.html)"
            )
            
            if filename:
                # Create a full report with both visualizations
                self.visualizer.save_full_report(filename)
                QMessageBox.information(self, "Success", f"Report saved to {filename}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not save visualization: {str(e)}")
    
    def closeEvent(self, event):
        """Handle window close event"""
        # Stop monitoring if active
        self.stop_monitoring()
        
        # Remove temporary files
        temp_files = ["temp_sankey.html", "temp_table.html", "temp_visualization.html"]
        for temp_file in temp_files:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                    print(f"Removed temporary file: {temp_file}")
                except Exception as e:
                    print(f"Error removing temporary file {temp_file}: {e}")
                    
        event.accept()

def main():
    app = QApplication(sys.argv)
    window = FilterTrailApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()