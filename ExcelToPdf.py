# Author: Elyece Malnati

import os
import win32com.client as win32
import configparser
import logging
from tqdm import tqdm

issues_to_print = []

# Set up logging
with open("log.txt", "w") as log_file:
    log_file.write("")

logging.basicConfig(filename="log.txt", level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
log = logging.getLogger()
log.debug("Starting program")

# Create or read a config file
default_config = {
    "working_directory": r"./",
    "sheets_to_print": ["Sheet 1", "Sheet 2"]
}

if not os.path.exists("config.ini"):
    with open("config.ini", "w") as config_file:
        configparser.ConfigParser(default_config).write(config_file)

config = configparser.ConfigParser(default_config)
config.read("config.ini")

# Validate the config file
working_directory = config.get("DEFAULT", "working_directory")
if working_directory == "":
    # default to current working directory
    working_directory = os.getcwd()
elif not os.path.exists(working_directory):
    log.error(f"The working directory '{working_directory}' does not exist. Please check the config file.")
    raise ValueError(f"The working directory '{working_directory}' does not exist. Please check the config file.")
try:
    sheets_to_print = config.get("DEFAULT", "sheets_to_print").strip("[]").split(",")
    # first trim the whitespace to account for user optionally including one or more
    sheets_to_print = [sheet_name.strip(" ") for sheet_name in sheets_to_print]
    # then strip the quotes
    sheets_to_print = [sheet_name.strip("'") for sheet_name in sheets_to_print]

except Exception as e:
    log.error(f"The 'sheets_to_print' option could not be read properly. Please check the config file.\n{e}")
    raise ValueError(f"The 'sheets_to_print' option could not be read properly. Please check the config file.\n{e}")

if not sheets_to_print or sheets_to_print == ['']:
    log.error("The 'sheets_to_print' option is empty. Please check the config file.")
    input("\nThe 'sheets_to_print' option is empty. Please check the config file and restart the program. " +
          "Press enter to exit ")
    raise ValueError("The 'sheets_to_print' option is empty. Exiting.")

# Get the file names from the working directory
file_names = [file_name for file_name in os.listdir(working_directory) if file_name.endswith(".xlsx")]
if not file_names:
    log.error("No Excel files found in the working directory. Please check the config file.")
    input(f"\nNo Excel files found in the working directory: {working_directory}. "
          "\nPlease check the config file and restart the program. Press enter to exit ")
    raise ValueError("No Excel files found in the working directory. Exiting.")

# Initialize Excel
excel = win32.Dispatch('Excel.Application')
excel.Visible = False

for file_name in tqdm(file_names):
    # Open the workbook
    workbook = excel.Workbooks.Open(os.path.join(working_directory, file_name))
    workbook_sheet_names = [sheet.Name for sheet in workbook.Sheets]
    for sheet_name in sheets_to_print:
        if sheet_name not in workbook_sheet_names:
            issue = f"Sheet '{sheet_name}' does not exist in workbook '{file_name}'. Please check the config file."
            issues_to_print.append(issue)
            log.error(issue)
            continue
        try:
            # Select the sheet
            sheet = workbook.Sheets(sheet_name)

            # Define the output PDF file name
            pdf_file_name = os.path.join(working_directory,
                                         f"{file_name.replace('.xlsx', '')} - {sheet_name}.pdf")

            # Print the sheet to PDF
            sheet.ExportAsFixedFormat(0, pdf_file_name)
            log.info(f"Sheet '{sheet_name}' has been saved as PDF: {os.path.abspath(pdf_file_name)}")
        except Exception as e:
            issue = f"Error printing sheet '{sheet_name}' in workbook '{file_name}': {str(e)}"
            log.error(issue)

    # Close the workbook
    workbook.Close(SaveChanges=False)

excel.Quit()
if issues_to_print:
    print("\nThe following issues were encountered:")
for issue in issues_to_print:
    print(issue)

input("All done, close the program or press enter to exit ")
