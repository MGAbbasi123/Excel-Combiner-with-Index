Excel Combiner with Index
This Python script automates the process of combining multiple Excel files spread across different subdirectories into a single, organized Excel workbook. It also generates an interactive "Home" sheet that acts as an index, allowing for easy navigation between the combined data sheets.

✨ Features
Batch Processing: Automatically scans a specified root directory for subfolders and processes all Excel files (.xls or .xlsx) found within them.

Data Consolidation: Concatenates data from all Excel files within a given subfolder into a single DataFrame, creating a dedicated sheet for each subfolder in the output workbook.

Source File Tracking: Adds a SourceFile column to each combined DataFrame, indicating the original Excel file from which each row of data originated.

Interactive Index: Creates a "Home" sheet with hyperlinks to each generated data sheet, enabling quick navigation.

Back-to-Home Links: Each data sheet includes a "← Back to Home" link for seamless return to the index.

Dynamic Sheet Naming: Uses subfolder names for sheet titles, truncating them to comply with Excel's 31-character limit.

Error Handling: Includes basic error handling for unreadable Excel files.

🛠️ Prerequisites
Before running the script, ensure you have the following Python libraries installed:

pandas

openpyxl

You can install them using pip:

pip install pandas openpyxl

🚀 Usage
Clone the Repository (or Download the Script):

git clone https://github.com/your_username/your_repo_name.git
cd your_repo_name

(Replace your_username and your_repo_name with your actual GitHub details.)

Configure the Root Directory:
Open the combine_excel.py (or whatever you name your script) file and modify the root_dir variable to point to the directory containing your subfolders with Excel files:

# Set the root directory
root_dir = Path("root_dir") # <-- Change this to your desired root directory

Important: Ensure the path uses forward slashes (/) or escaped backslashes (\\) for cross-platform compatibility.

Run the Script:
Execute the script from your terminal:

python combine_excel.py

📂 File Structure
The script expects your Excel files to be organized within subfolders inside the root_dir:

(root_dir)/ 
    ├── Folder A/
    │   ├── data1.xls
    │   └── data2.xlsx
    ├── Folder B/
    │   ├── sales.xlsx
    │   └── purchases.xls
    └── Folder C/
        └── report.xlsx

📊 Output
The script will generate a new Excel file named Combined_Workbook_with_Index.xlsx in your specified root_dir. This workbook will contain:

A Home sheet with clickable links to each combined data sheet.

A dedicated sheet for each subfolder (e.g., "Folder A", "Folder B", "Folder C"), containing all the data from Excel files within that respective subfolder, plus a SourceFile column.

A success message will be printed to the console upon completion.

✅ Workbook created successfully: (root_dir)\(file_name).xlsx
