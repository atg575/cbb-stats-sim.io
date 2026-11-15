import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter
from openpyxl.formula.translate import Translator
from copy import copy
import os
import re

def fallback_translate(formula, old_row, new_row):
    """A safe regex fallback for dragging down formulas that openpyxl's Translator cannot handle."""
    if hasattr(formula, 'text'):
        formula_str = formula.text
    elif not isinstance(formula, str):
        formula_str = str(formula)
    else:
        formula_str = formula

    def replacer(match):
        col = match.group(1)
        row_str = match.group(2)
        if row_str == str(old_row):
            return f"{col}{new_row}"
        return match.group(0)
    
    return re.sub(r'([A-Za-z]{1,3})(\d+)', replacer, formula_str)

def add_row_to_excel_tables(file_path, target_tables):
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        return

    print(f"Opening {file_path}...")
    
    wb = openpyxl.load_workbook(file_path, keep_vba=False)
    tables_updated = 0

    # Loop through every sheet in the workbook
    for sheet in wb.worksheets:
        
        # Loop through the actual Table objects directly
        for table in sheet.tables.values():
            
            # Check if the table matches your target list (case-insensitive)
            if table.name and table.name.lower() in [t.lower() for t in target_tables]:
                
                # Get the current boundaries of the table
                min_col, min_row, max_col, max_row = range_boundaries(table.ref)
                old_last_row = max_row
                new_row = max_row + 1
                
                # Create the new reference string expanded down by 1 row
                new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_row}"
                table.ref = new_ref
                
                # Copy formulas and formatting from the old last row to the new one
                for col in range(min_col, max_col + 1):
                    old_cell = sheet.cell(row=old_last_row, column=col)
                    new_cell = sheet.cell(row=new_row, column=col)
                    
                    # If the cell above is a formula, translate it down one row
                    if old_cell.data_type == 'f':
                        try:
                            # Try the official openpyxl translator first
                            new_formula = Translator(old_cell.value, origin=old_cell.coordinate).translate_formula(new_cell.coordinate)
                        except TypeError:
                            # If it hits an array formula bug, use our custom regex tool
                            new_formula = fallback_translate(old_cell.value, old_last_row, new_row)
                        
                        new_cell.value = new_formula
                    else:
                        new_cell.value = None # Leave non-formula cells blank for your new data
                        
                    # Copy styling (borders, fonts, number formats) so the table looks continuous
                    if old_cell.has_style:
                        new_cell.font = copy(old_cell.font)
                        new_cell.border = copy(old_cell.border)
                        new_cell.fill = copy(old_cell.fill)
                        new_cell.number_format = copy(old_cell.number_format)
                        new_cell.alignment = copy(old_cell.alignment)
                
                print(f"Success: Added row {new_row} (with formulas) to '{table.name}' on '{sheet.title}'.")
                tables_updated += 1
                
    if tables_updated > 0:
        print("Saving workbook...")
        wb.save(file_path)
        print("Done!")
    else:
        print("No matching tables were found in the workbook.")
        wb.close()

if __name__ == "__main__":
    file_name = "Batch_Simulation_Results.xlsx" 
    tables_to_update = ["BaseWeek", "BaseWeek11", "BaseWeek13", "BaseWeek15", "BaseWeek17", "BaseWeek19","ninetyplus", "ninetyplus21", "ninetyplus22", "ninetyplus23", "ninetyplus24", "ninetyplus25", "ninetyplus26", "ninetyplus27", "ninetyplus28"] 
    add_row_to_excel_tables(file_name, tables_to_update)

