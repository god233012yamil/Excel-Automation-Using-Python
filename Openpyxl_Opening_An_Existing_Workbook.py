import openpyxl
from openpyxl import Workbook

def open_existing_workbook(file_name: str) -> None:
    """
    Opens an existing workbook, reads data from the active 
    sheet, and prints it.

    :param file_name: Name of the existing Excel file to open.
    """
    try:
        # Load the existing workbook
        wb: Workbook = openpyxl.load_workbook(file_name)

        # Get the active worksheet
        ws = wb.active

        # Print the title of the active worksheet
        print(f"Worksheet title: {ws.title}")

        # Read and print the data
        for row in ws.iter_rows(values_only=True):
            print(row)

        print("Workbook loaded and data read successfully.")

    except FileNotFoundError as e:
        print(f"Error: {e}")
        print("The specified file does not exist.")

# Specify the file name of the existing workbook
file_name = "example.xlsx"

# Open the existing workbook and read data
open_existing_workbook(file_name)
