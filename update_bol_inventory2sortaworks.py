import os
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def load_excel_file(file_path, sheet_name, header_row):
    # Load without altering column case
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
    df.columns = df.columns.str.strip()  # Only strip whitespace, retain case
    return df

def update_arrow(arrow_file, wp_file):
    # Load Arrow data with header on Row 9
    arrow_df = load_excel_file(arrow_file, "Transaction Entry", header_row=8)
    wp_df = pd.read_excel(wp_file)
    wp_df.columns = wp_df.columns.str.strip()  # Only strip whitespace

    # Print processed columns for verification
    print("Processed Arrow file columns:", list(arrow_df.columns))
    print("Processed WP file columns:", list(wp_df.columns))

    # Define mappings without lowercase standardization
    mappings = {
        'Settle No': 'Settle #',
        'Vehicle Id': 'BOL',  # Correct mapping based on WP file data
        'Applied': 'Applied',
        'Contract No': 'Contract #'
    }

    # Check if required columns exist
    for wp_field, arrow_field in mappings.items():
        if wp_field not in wp_df.columns:
            print(f"Missing WP column: '{wp_field}'")
            return
        if arrow_field not in arrow_df.columns:
            print(f"Missing Arrow column: '{arrow_field}'")
            return

    # Ensure columns are strings for proper comparison
    for wp_field, arrow_field in mappings.items():
        wp_df[wp_field] = wp_df[wp_field].astype(str)
        arrow_df[arrow_field] = arrow_df[arrow_field].astype(str)

    # Iterate over WP rows to find matching BOLs in Arrow
    for index, wp_row in wp_df.iterrows():
        try:
            wp_bol = wp_row[mappings['Vehicle Id']]  # Use 'Vehicle Id' for WP
            wp_settle = wp_row[mappings['Settle No']]
            wp_contract = wp_row[mappings['Contract No']]
            wp_applied = wp_row[mappings['Applied']]

            # Debug print to confirm retrieved values
            print(f"Processing WP Row {index}: BOL={wp_bol}, Settle={wp_settle}, Contract={wp_contract}, Applied={wp_applied}")

            # Find matching rows in Arrow by BOL
            match_rows = arrow_df[arrow_df['BOL'] == wp_bol]  # Match using original case

            if not match_rows.empty:
                total_applied = match_rows[mappings['Applied']].astype(float).sum()
                if float(wp_applied) != total_applied:
                    base_row_index = match_rows.index[0]
                    arrow_df.at[base_row_index, mappings['Applied']] = wp_applied
                    arrow_df.at[base_row_index, mappings['Settle No']] = wp_settle
                    arrow_df.at[base_row_index, mappings['Contract No']] = wp_contract

                    for i, (_, row) in enumerate(match_rows[1:].iterrows(), start=1):
                        new_row = row.copy()
                        new_row[mappings['Applied']] = float(wp_applied) / (len(match_rows) + 1)
                        new_row[mappings['Contract No']] = wp_contract
                        new_row[mappings['Settle No']] = wp_settle

                        # Append new row to Arrow DataFrame
                        arrow_df = pd.concat([arrow_df, pd.DataFrame([new_row])], ignore_index=True)

            else:
                print(f"BOL {wp_bol} not found in Arrow file.")
        except KeyError as e:
            print(f"Error accessing column in WP Row {index}: {e}")

    # Save the updated Arrow file
    updated_path = os.path.join(os.path.dirname(arrow_file), "updated_inventory.xlsx")
    arrow_df.to_excel(updated_path, index=False)
    print(f"Updated Arrow inventory saved as {updated_path}")

def main():
    # Use Tkinter to open file dialogs
    root = Tk()
    root.withdraw()

    # Ask for Arrow and WP files
    arrow_file = askopenfilename(title="Select the Arrow inventory file", filetypes=[("Excel files", "*.xlsx")])
    if not arrow_file:
        print("Arrow file not selected. Exiting.")
        return

    wp_file = askopenfilename(title="Select the WP data dump file", filetypes=[("Excel files", "*.xlsx")])
    if not wp_file:
        print("WP file not selected. Exiting.")
        return

    update_arrow(arrow_file, wp_file)

    # Keep the command window open
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()




