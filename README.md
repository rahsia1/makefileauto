# makefileauto
Excel automation
# Work Instruction Update Tool

This VBA script updates a "Work Instruction" sheet based on user input from a "Part Information" sheet. It copies relevant data from the chosen part row to a new sheet in a new workbook and saves it with a generated name.

## Usage

1. Open the Excel workbook containing the "Part Information" and "Work Instruction" sheets.
2. Run the VBA script by clicking on the "Run Macro" button.
3. Enter the row number of the part you want to copy when prompted.
4. The script copies relevant data to a new workbook and saves it.

## Code Explanation

The VBA script performs the following tasks:
- Asks the user for the row number of the part to copy.
- Validates the input and checks for cancellation.
- Copies specific data from the "Part Information" sheet to the "Work Instruction" sheet.
- Generates a new name based on the values in the chosen row.
- Creates a new workbook, copies the "Work Instruction" sheet, and saves it with the generated name.

## Modification

You can modify the script as needed, such as adjusting the cell references or updating the save location.

## How to Run

To run the script, follow these steps:
1. Press `Alt + F8` to open the "Macro" dialog.
2. Select "UpdateWiOnButtonClick" and click "Run."

Feel free to customize and improve the script based on your requirements.
