#!/usr/bin/env python3
"""
Example of reading and writing Excel files using openpyxl
"""

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime


def write_excel_example():
    """Create and write data to an Excel file"""
    # Create a new workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    
    # Add headers
    headers = ["Name", "Age", "City", "Date"]
    ws.append(headers)
    
    # Style headers
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    # Add data rows
    data = [
        ["Alice", 30, "New York", datetime.now()],
        ["Bob", 25, "San Francisco", datetime.now()],
        ["Charlie", 35, "Boston", datetime.now()],
    ]
    
    for row in data:
        ws.append(row)
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 20
    
    # Save the workbook
    wb.save('example.xlsx')
    print("✓ Excel file 'example.xlsx' created successfully")


def read_excel_example():
    """Read data from an Excel file"""
    try:
        # Load the workbook
        wb = load_workbook('example.xlsx')
        ws = wb.active
        
        print("\n✓ Reading from 'example.xlsx':")
        print("-" * 60)
        
        # Read and display all rows
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
            print(row)
        
        print("-" * 60)
        
        # Access specific cells
        print(f"\nFirst person's name: {ws['A2'].value}")
        print(f"Number of data rows: {ws.max_row - 1}")
        
    except FileNotFoundError:
        print("Error: 'example.xlsx' not found. Please run write_excel_example() first.")


def update_excel_example():
    """Update existing Excel file"""
    try:
        # Load existing workbook
        wb = load_workbook('example.xlsx')
        ws = wb.active
        
        # Add a new row
        ws.append(["David", 28, "Chicago", datetime.now()])
        
        # Modify a cell
        ws['C2'] = "Los Angeles"  # Change Bob's city
        
        # Save changes
        wb.save('example.xlsx')
        print("\n✓ Excel file updated successfully")
        
    except FileNotFoundError:
        print("Error: 'example.xlsx' not found.")


if __name__ == "__main__":
    write_excel_example()
    read_excel_example()
    update_excel_example()
    read_excel_example()  # Show updated data
