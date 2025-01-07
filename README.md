# SAS to Excel Converter - Quick Start Guide

This guide will help you set up and use the SAS to Excel converter script. The script converts .sas7bdat and .xpt files to Excel format (.xlsx).

## Setup Instructions

### 1. Install Python
1. Download Python 3.11 from https://www.python.org/downloads/
2. Run the installer
3. **IMPORTANT:** Check the box "Add Python to PATH" during installation
   - This is a crucial step - the script won't work if you miss it!

### 2. Get the Script
1. Download the script from: https://github.com/FIECON/Centralised-code-store/tree/main/File%20type%20converters
2. Save the file `sas_converter.py` to your computer
3. Make note of where you saved it

### 3. Install Required Packages
1. Open Command Prompt
   - Press Windows key + R
   - Type "cmd" and press Enter
2. Type these commands one by one:
   ```
   pip install pandas
   pip install pyreadstat
   pip install openpyxl
   ```

## Using the Converter

### Basic Usage
1. Put all your SAS files (.sas7bdat and .xpt) in one folder
2. Open Command Prompt
3. Navigate to where you saved the script, for example:
   ```
   cd C:\Users\YourName\Desktop
   ```
4. Run the script:
   ```
   python sas_converter.py "path\to\your\folder"
   ```
5. Find your converted Excel files in the same folder
   - Files will be named like: "filename-xpt.xlsx" or "filename-sas7bdat.xlsx"

### Advanced Usage
- To save files in a different folder:
  ```
  python sas_converter.py "path\to\your\folder" --output "path\to\output\folder"
  ```

## Troubleshooting

### Common Issues and Solutions

1. "Python/pip is not recognized"
   - Solution: Restart Command Prompt after installation
   - If still not working: Reinstall Python and make sure to check "Add Python to PATH"

2. "Module not found" errors
   - Solution: Make sure you completed step 3 (Install Required Packages)
   - Try running the pip install commands again

3. Can't find converted files
   - They're in the same folder as your SAS files by default
   - Check the command window for any error messages
   - Make sure your input path is correct

4. Permission errors
   - Try running Command Prompt as administrator
   - Check if you have write permissions in the folder

### Need More Help?
Contact Damian Kosinski (damian.kosinski@fiecon.com)
