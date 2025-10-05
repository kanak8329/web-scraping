import pandas as pd


def update_rankings_in_place(main_file_path, source_file_path):
    """
    Updates the main QS rankings file in-place with data from a source file.

    It matches institutions by name and fills in any missing data for total
    students, international students, total faculty, and international faculty
    percentage. The original main file is overwritten with the changes.

    Args:
        main_file_path (str): Path to the main QS rankings Excel file to be updated.
        source_file_path (str): Path to the source Excel file with the data.
    """
    try:
        # Load the main rankings file. The actual headers are on the 3rd row (index 2).
        df_main = pd.read_excel(main_file_path, header=2)
        print(f"Successfully loaded '{main_file_path}' with {len(df_main)} rows.")

        # Load the source data file.
        df_source = pd.read_excel(source_file_path)
        print(f"Successfully loaded '{source_file_path}' with {len(df_source)} rows.")

        # --- Column Mapping and Renaming ---
        # The main file's header is read from its 3rd row, resulting in generic
        # column names. We map the source columns to these target column names.
        column_mapping = {
            'Institution': 'Name',
            'Total Students': 'Column1',
            'International Students': 'Column3',
            'Total Faculty Staff': 'staff',
            "Int'l Staff %": 'Column2'
        }
        df_source_renamed = df_source.rename(columns=column_mapping)

        # --- Data Cleaning ---
        # Strip whitespace from institution names to ensure accurate matching.
        df_main['Name'] = df_main['Name'].str.strip()
        df_source_renamed['Name'] = df_source_renamed['Name'].str.strip()

        # Clean source data: convert numeric strings (e.g., "1,234") to numbers.
        numeric_cols = ['Column1', 'Column3', 'staff']
        for col in numeric_cols:
            if col in df_source_renamed.columns:
                df_source_renamed[col] = pd.to_numeric(
                    df_source_renamed[col].astype(str).str.replace(',', ''),
                    errors='coerce'
                )

        # Clean percentage column: remove '%' and convert to a number.
        if 'Column2' in df_source_renamed.columns:
            df_source_renamed['Column2'] = pd.to_numeric(
                df_source_renamed['Column2'].astype(str).str.replace('%', ''),
                errors='coerce'
            )

        # --- Update Main DataFrame ---
        # Set 'Name' as the index for both DataFrames to align rows for updating.
        df_main.set_index('Name', inplace=True)
        df_source_renamed.set_index('Name', inplace=True)

        # The columns we intend to potentially fill in the main dataframe.
        update_cols = ['staff', 'Column1', 'Column2', 'Column3']

        print(f"Checking for and filling missing data in columns: {update_cols}")

        # update() with overwrite=False fills NaN (blank) values in df_main
        # without overwriting existing data.
        df_main.update(df_source_renamed[update_cols], overwrite=False)

        # Reset the index to turn 'Name' back into a column.
        df_main.reset_index(inplace=True)

        print("Update process complete.")

        # --- Saving the result ---
        # Overwrite the original main file with the updated data.
        df_main.to_excel(main_file_path, index=False)
        print(f"Successfully saved changes back to '{main_file_path}'")

    except FileNotFoundError as e:
        print(f"Error: {e}. Please ensure the file paths are correct.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


if __name__ == '__main__':
    # Define the input file paths.
    # The main_file will be read and then overwritten with the updates.
    main_file = 'Copy of 2026_QS_World_University_Rankings_1.2_(For_qs.com)_(3)(1).xlsx'
    source_file = 'qs_table_827.xlsx'

    update_rankings_in_place(main_file, source_file)