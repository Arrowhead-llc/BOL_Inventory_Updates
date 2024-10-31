import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename

def load_excel_file(prompt):
    root = Tk()
    root.withdraw()
    file_path = askopenfilename(title=prompt, filetypes=[("Excel files", "*.xlsx")])
    return file_path if file_path else None

def update_arrow(arrow_file, wp_file):
    # Load data from the Arrow and WP files
    arrow_wb = openpyxl.load_workbook(arrow_file)
    arrow_sheet = arrow_wb['Transaction Entry']
    arrow_df = pd.DataFrame(arrow_sheet.values)
    arrow_df.columns = arrow_df.iloc[0]  # Use the first row as the header
    arrow_df = arrow_df[1:]  # Skip the header row

    # Display the columns in the Arrow sheet
    print("Arrow file columns:", arrow_df.columns.tolist())

    wp_df = pd.read_excel(wp_file)

    # Display the columns in the WP file
    print("WP file columns:", wp_df.columns.tolist())

    # Define column mappings
    mappings = {
        'Settle No': 'Settle #',
        'Vehicle Id': 'BOL',
        'Applied': 'Applied',
        'Contract No': 'Contract #'
    }

    # Ensure all relevant columns are present and formatted as strings
    for wp_col, arrow_col in mappings.items():
        if wp_col in wp_df.columns:
            wp_df[wp_col] = wp_df[wp_col].astype(str)
        if arrow_col in arrow_df.columns:
            arrow_df[arrow_col] = arrow_df[arrow_col].astype(str)
    
    split_rows = []

    # Iterate through WP Data and match with Arrow by BOL and Settle #
    for _, wp_row in wp_df.iterrows():
        wp_bol = wp_row[mappings['Vehicle Id']]
        wp_settle_no = wp_row[mappings['Settle No']]
        matching_rows = arrow_df[(arrow_df['BOL'] == wp_bol) & (arrow_df['Settle #'] == wp_settle_no)]
        
        if not matching_rows.empty:
            for idx, match_row in matching_rows.iterrows():
                count_value = float(match_row['Count']) if '.' not in str(match_row['Count']) else float(match_row['Count'].split('.')[0])
                # Add first entry or increment for new splits
                if wp_row[mappings['Contract No']] == match_row['Contract #'] and wp_row[mappings['Applied']] == match_row['Applied']:
                    continue  # Skip identical records

                # Create new row with incremented count and relevant fields
                new_row = match_row.copy()
                new_count = f"{int(count_value)}.{len(split_rows) + 1}"  # Add decimal for splits
                new_row['Count'] = new_count
                new_row['Contract #'] = wp_row[mappings['Contract No']]
                new_row['Applied'] = wp_row[mappings['Applied']]
                split_rows.append(new_row)

    # Append split rows to the original DataFrame
    updated_df = pd.concat([arrow_df, pd.DataFrame(split_rows)]).sort_values(by=['Count']).reset_index(drop=True)

    # Update the Excel sheet with the modified data
    for row_idx, row_data in enumerate(dataframe_to_rows(updated_df, index=False, header=True), start=1):
        for col_idx, cell_value in enumerate(row_data, start=1):
            arrow_sheet.cell(row=row_idx, column=col_idx, value=cell_value)

    # Save to a new file
    save_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                  title="Save the updated Arrow inventory")
    if save_path:
        arrow_wb.save(save_path)
        print(f"Updated Arrow inventory saved as {save_path}")
    
    input("Press Enter to exit...")

def main():
    print("Select the Arrow inventory file:")
    arrow_file = load_excel_file("Select the Arrow inventory file")
    if not arrow_file:
        print("Arrow file not selected. Exiting...")
        return

    print("Select the WP data dump file:")
    wp_file = load_excel_file("Select the WP data dump file")
    if not wp_file:
        print("WP file not selected. Exiting...")
        return

    update_arrow(arrow_file, wp_file)

if __name__ == "__main__":
    main()
input("Press Enter to exit...")