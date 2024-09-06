ExcelToPdf is a simple tool to iterate through Excel files in a folder and automatically save selected worksheets as pdf files. Every sheet will be saved in the same folder as the Excel file, as its own PDF.
The program reads the config file 'config.ini', located in the same directory as the executable, to determine the working directory and the worksheets to print. If a config file does not exist, it creates one with default values during the first run.
Every worksheet in the config needs to be entered as a list element in single quotes, e.g. sheets_to_print = ['Sheet 1', 'Sheet 2']


Program icon was created by Freepik