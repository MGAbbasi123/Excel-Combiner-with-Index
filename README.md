Excel Combiner Desktop App
This is a user-friendly desktop application built with Python's Tkinter library that automates the process of combining multiple Excel files from various subdirectories into a single, organized Excel workbook. It also creates an interactive "Home" sheet for easy navigation.

This application is designed to run directly on your local computer and allows you to select a root directory from your file system.

✨ Features
Local Folder Selection: Easily select any root directory on your local computer using a familiar file browser interface.

Batch Processing: Automatically scans the selected root directory for subfolders and processes all Excel files (.xls or .xlsx) found within them.

Data Consolidation: Combines data from all Excel files within a given subfolder into a single DataFrame, creating a dedicated sheet for each subfolder in the output workbook.

Source File Tracking: Adds a SourceFile column to each combined DataFrame, indicating the original Excel file from which each row of data originated.

Interactive Index: Generates a "Home" sheet with clickable hyperlinks to each created data sheet, enabling quick navigation within the output workbook.

Back-to-Home Links: Each data sheet includes a "← Back to Home" link for seamless return to the index.

Dynamic Sheet Naming: Uses subfolder names for sheet titles, truncating them to comply with Excel's 31-character limit.

User Feedback: Provides real-time status messages and pop-up alerts for successful operations or errors.

🛠️ Prerequisites
To run this application, you need Python installed on your computer, along with the following libraries:

pandas

openpyxl

tkinter (This is usually included with standard Python installations)

You can install the required libraries using pip from your terminal or command prompt:

pip install pandas openpyxl

🚀 How to Run the Application
Save the Script:
Copy the entire Python code for the Tkinter application (the one with import tkinter as tk) and save it to a file named excel_combiner_app.py (or any .py name you prefer) on your local computer.

Open Your Terminal/Command Prompt:
Navigate to the directory where you saved excel_combiner_app.py.

cd /path/to/your/script/directory

Execute the Script:
Run the script using your Python interpreter:

python excel_combiner_app.py

A new desktop window titled "Excel Combiner" should appear.

⚙️ How to Use the App
Select Root Directory:
Click the "Browse Folder" button. A file explorer window will open. Navigate to and select the main directory that contains the subfolders with your Excel files. The selected path will appear in the "Selected Root Directory" field.

Process Files:
Once a directory is selected, the "Process Excel Files" button will become active. Click it to start the combination process.

View Output:
Upon successful completion, a success message will appear, and the path to the newly created combined Excel workbook will be displayed in the "Output File" field. An informational pop-up will also confirm the success and the output path.

📂 Expected File Structure
The script expects your Excel files (.xls or .xlsx) to be organized within subfolders inside the root_dir you select.

Your_Selected_Root_Directory/
├── Project_A/
│   ├── data_q1.xls
│   └── data_q2.xlsx
├── Department_B/
│   ├── sales_report.xlsx
│   └── inventory.xls
└── Category_C/
    └── summary.xlsx

📊 Output
The application will generate a new Excel file (e.g., Combined_Workbook_with_Index_YYYYMMDD_HHMMSS.xlsx) directly within the root_dir you selected. This workbook will contain:

A Home sheet serving as an index with clickable links to each combined data sheet.

A dedicated sheet for each subfolder (e.g., "Project A", "Department B", "Category C"), containing all the data from Excel files within that respective subfolder, plus a SourceFile column.

💡 Troubleshooting
_tkinter.TclError: no display name and no $DISPLAY environment variable
This error occurs when you try to run a graphical application (like this Tkinter app) in an environment without a display server. This is common if you're trying to run it on a remote server, a virtual machine without a desktop environment, or within certain online coding platforms.

Solution: This application is designed to run directly on your local computer where you have a graphical user interface. Ensure you are running the excel_combiner_app.py script from your local machine's terminal or command prompt, not on a remote server or a headless environment.