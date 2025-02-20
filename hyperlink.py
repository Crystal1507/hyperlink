import openpyxl

# Path to your Excel file
file_path = "/Users/shinyjoycesj/Desktop/CoNames.xlsx"  # Replace with your actual file path

# Load the Excel workbook
wb = openpyxl.load_workbook(file_path)

# Select the active sheet (or use wb['SheetName'] for a specific sheet)
sheet = wb.active

# Extract hyperlinks from column A and write to column B
for row in range(2, sheet.max_row + 1):  # Assuming row 1 is headers
    cell = sheet[f"A{row}"]
    if cell.hyperlink:  # Check if the cell contains a hyperlink
        # Write the hyperlink target (URL) into column B
        sheet[f"B{row}"].value = cell.hyperlink.target
    else:
        # No hyperlink, write "No URL" or leave it blank
        sheet[f"B{row}"].value = "No URL"

# Save the updated workbook
output_file = "/Users/shinyjoycesj/Desktop/CoNames_with_URLs.xlsx"  # Output file path
wb.save(output_file)

print(f"URLs have been written to column B and saved in '{output_file}'.")
