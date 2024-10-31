import pandas as pd
from tkinter import Tk, filedialog
import os

def select_file():
    Tk().withdraw()  # Hides the main Tkinter window
    file_path = filedialog.askopenfilename(title="Select the Excel file")
    return file_path

def update_inventory(inventory_path, wp_path):
    # Load Arrow inventory with headers starting from Row 9
    inventory = pd.read_excel(inventory_path, sheet_name="Transaction Entry", skiprows=8)
    print("Original Arrow file columns:", inventory.columns)

    # Load WP data dump
    wp_data = pd.read_excel(wp_path)
    print("Original WP file columns:", wp_data.columns)

    # Standardize column names for matching
    inventory.columns = inventory.columns.str.strip().str.replace(' ', '_').str.replace('#', 'Num')
    wp_data.columns = wp_data.columns.str.strip().str.replace(' ', '_').str.replace('#', 'Num')

    # Process WP data and update inventory
    for _, wp_row in wp_data.iterrows():
        # Find rows in Arrow matching the BOL and Settle #
        matched_rows = inventory[(inventory["Settle_Num"] == wp_row["Settle_No"]) & 
                                 (inventory["BOL"] == wp_row["Vehicle_Id"])]

        # Handle cases where BOL has multiple contract splits
        if not matched_rows.empty:
            initial_index = matched_rows.index[0]
            # Update the first matched row with the WP data
            inventory.loc[initial_index, "Applied"] = wp_row["Applied"]
            inventory.loc[initial_index, "Contract_Num"] = wp_row["Contract_No"]

            # Add additional lines if more than one contract is applied
            if len(matched_rows) > 1:
                for i, row_index in enumerate(matched_rows.index[1:], start=1):
                    split_row = inventory.loc[row_index].copy()
                    split_row["Applied"] = wp_row["Applied"]
                    split_row["Contract_Num"] = wp_row["Contract_No"]
                    split_row["Count"] = f"{split_row['Count']}.{i}"
                    inventory = pd.concat([inventory.iloc[:row_index + i], pd.DataFrame([split_row]), inventory.iloc[row_index + i:]], ignore_index=True)
        else:
            # Add new row if BOL is not found in Arrow
            new_row = {
                "BOL": wp_row["Vehicle_Id"],
                "Settle_Num": wp_row["Settle_No"],
                "Applied": wp_row["Applied"],
                "Contract_Num": wp_row["Contract_No"]
            }
            inventory = pd.concat([inventory, pd.DataFrame([new_row])], ignore_index=True)
    
    # Save updated inventory
    updated_inventory_path = os.path.join(
        os.path.dirname(inventory_path),
        "test_2023_corn_database_style_updated_with_splits.xlsx"
    )
    inventory.to_excel(updated_inventory_path, index=False)
    print(f"Updated Arrow inventory saved as {updated_inventory_path}")

def main():
    print("Select the Arrow inventory file:")
    inventory_path = select_file()
    
    print("Select the WP data dump file:")
    wp_path = select_file()
    
    update_inventory(inventory_path, wp_path)

if __name__ == "__main__":
    main()








#if __name__ == "__main__":
#    main()

