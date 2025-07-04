# FilterTrail


https://www.youtube.com/watch?v=B2SaTRrf5Hs
A lightweight Windows application that monitors Excel filter operations, records filter sequences and data counts, and visualizes the filtration process using Sankey diagrams.

![image](https://github.com/user-attachments/assets/05d32b4c-3e63-4eb0-b6ce-5452397beaeb)


## Features
- Real-time monitoring of Excel filter operations
- Tracking of filter sequence and filter steps
- Counting of data rows after each filter operation using primary key column
- Live visualization of filter flow with Sankey diagrams
- Ability to edit filter step names after monitoring stops
- Synchronization between table edits and the Sankey diagram

## System Requirements
- Windows operating system (7/8/10/11)
- Microsoft Excel (2013 or newer)
- No Python installation required for the executable version

## Installation Options

### Option 1: Use the executable (recommended)
1. Download the FilterTrail.exe file from the releases section
2. Run the executable directly - no installation needed

### Option 2: Run from source code
1. Make sure you have Python 3.8+ installed on your Windows machine
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Run the application:
   ```
   python filter_trail.py
   ```

## Usage Guide
1. Open your Excel file and select the sheet you want to monitor
2. Start FilterTrail
3. (Optional) Select the primary key column (default is Column 1)
4. Click "Start Monitoring" button
5. Apply filters to your Excel data
6. Watch the Sankey diagram and table update in real-time, showing your filter sequence
7. When done, click "Stop" to end monitoring
8. After stopping, you can edit step names in the table by:
   - Double-clicking on any step name in the table
   - Using the "Edit" button in the table
9. Changes in step names will automatically update in the Sankey diagram
10. Click "Save" to export the Sankey diagram as an HTML file

## Key Features Explained

### Primary Key Column
- Counts non-empty cells in the designated column when connecting to Excel
- Provides more accurate counting compared to using all rows in the sheet
- Can be changed before starting monitoring

### Filter Column Renaming
- Global renaming: Use the "Rename Filter Column" section to rename a column throughout all steps
- Step-specific renaming: After stopping monitoring, double-click a step name in the table to edit just that step

### Visualization Options
- Table view: Shows filter steps, timestamps, row counts, and filter values
- Sankey diagram: Visual representation of data flow through filter steps
- Both views stay synchronized when editing step names

## Building from Source

### Standard Build (single file)
```
python build_scripts/optimized_build.py
```

### Anaconda Environment Build
If you're using Anaconda and encounter issues with the standard build:
```
python build_scripts/override_build.py
```

The executable will be created in the `dist` folder.

## Project Structure
- `main.py` - Main application code
- `filter_trail.py` - Application launcher
- `requirements.txt` - Python dependencies
- `filter_data.json` - Initial empty data file
- `reset_data.py` - Utility to reset filter data
- `build_scripts/` - Python build scripts
- `batch_files/` - Windows batch files (for Windows users)

## Troubleshooting
- If the application can't connect to Excel, make sure Excel is open with a workbook
- If step names don't update in the Sankey diagram, try clicking "Refresh" button
- For large Excel files (>100,000 rows), the monitoring might be slower
- If you encounter any issues with filter detection, try applying filters one at a time
- When saving visualization to HTML, ensure you have write permissions to the selected folder
