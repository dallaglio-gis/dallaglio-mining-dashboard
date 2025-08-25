INSTRUCTIONS: Running the CEO Daily Report Automation Script
============================================================

This folder contains the automation tools for generating daily CEO report sheets from the June 2025 production data.

FILES REQUIRED:
---------------
- Geology Daily Work Plan <MonthYear>.xlsx       (destination workbook)
- <Month>_<Year>_DAILY_REPORT.xlsx              (source data workbook, e.g. June_2025_DAILY_REPORT.xlsx, July_2025_DAILY_REPORT.xlsx)
- update_daily_report_all_days.py               (Python script to run)
- run_all_days.bat                              (optional Windows batch file)

IMPORTANT:
----------
- The script automatically detects the month and year from the source file name (e.g., July_2025_DAILY_REPORT.xlsx will process July 2025 data).
- To process a new month, simply provide a source file named in the format <Month>_<Year>_DAILY_REPORT.xlsx (e.g., August_2025_DAILY_REPORT.xlsx). No script changes are needed.
- The destination file can also be renamed to match the month if desired (e.g., Geology Daily Work Plan July2025.xlsx), but make sure to update the DEST_FILE variable in the script if you do so.

HOW TO USE:
-----------
1. Make sure BOTH Excel files (Geology Daily Work Plan June2025.xlsx and June_2025_DAILY_REPORT.xlsx) are CLOSED before running the script.

2. You have several ways to run the script:

   - **Open a command prompt in this folder and run:**
     python update_daily_report_all_days.py

   - **Double-click update_daily_report_all_days.py** (if Python is associated with .py files)

   - **Double-click run_all_days.bat**
     - This will open a Command Prompt window and automatically run the script for you.
     - The window will stay open so you can see the results (success messages, errors, etc.).
     
3. Wait for the script to finish. It will tell you if new daily sheets were created, or if everything is already up to date.

4. Open the destination file (e.g., Geology Daily Work Plan <MonthYear>.xlsx) to see the new daily report sheets (if any were added).

NOTES:
------
- The script will update existing daily sheets if source data changes (including previous days).
- MTD and Budget MTD values are now recalculated cumulatively for each day, using all data up to that date.
- Budget MTD uses daily budget values from the source file (not static values).
- If previous days' data is revised, all affected MTD and Budget MTD values are recalculated for subsequent days.
- Budget value row numbers are configurable in the script (see extract_daily_data_for_month function).
- See VERSION_NOTES.txt for a summary of recent changes and enhancements.
- If you add new data to the source file later, just rerun the script to generate any new daily sheets.
- To process a different month/year, just use a source file named <Month>_<Year>_DAILY_REPORT.xlsx (e.g., September_2025_DAILY_REPORT.xlsx). The script will automatically adapt.
- If you rename the destination file, update the DEST_FILE variable in update_daily_report_all_days.py accordingly.

TROUBLESHOOTING:
----------------
- If you get a PermissionError, make sure both Excel files are closed.
- If you get a Python error about missing modules, install openpyxl with:
    pip install openpyxl

For any issues or questions, contact your automation script maintainer.
