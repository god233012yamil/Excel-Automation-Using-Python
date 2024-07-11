import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side

def create_new_workbook(file_name: str) -> None:
    """
    Creates a new workbook with sample data and saves it to the 
    specified file.

    :param file_name: Name of the file to save the workbook as.
    """
    # Create a new workbook
    wb: Workbook = Workbook()
    
    # Get the active worksheet
    ws = wb.active
    ws.title = "Sample Data"

    # Add data to the worksheet
    headers = ["Name", "Age", "Department"]
    data = [
        ["John Doe", 30, "Sales"],
        ["Jane Smith", 25, "Marketing"],
        ["Emily Davis", 22, "Development"]
    ]
    
    # Append headers
    ws.append(headers)
    
    # Append data rows
    for row in data:
        ws.append(row)    

    # Save the workbook
    wb.save(file_name)
    print(f"Workbook '{file_name}' created and saved successfully.")

# Specify the file name
file_name = "new_workbook.xlsx"

# Create the new workbook
create_new_workbook(file_name)
