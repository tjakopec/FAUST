import pandas as pd
import os

def clean_dataframe(df):
    """
    Cleans the dataframe by:
    - Dropping unnecessary columns.
    - Renaming columns.
    - Removing 'Ukupno' rows.
    - Resetting the index.
    """
    # Columns to drop
    cols_to_drop = ['Unnamed: 2', 'SJEDIŠTE /\nPREBIVALIŠTE\nPRIMATELJA.1']
    
    # Drop columns that exist in the dataframe
    df = df.drop(columns=[col for col in cols_to_drop if col in df.columns])

    # Columns to rename
    rename_dict = {
        'NAZIV PRIMATELJA': 'NAZIV_PRIMATELJA',
        'OIB PRIMATELJA': 'OIB_PRIMATELJA',
        'SJEDIŠTE /\nPREBIVALIŠTE\nPRIMATELJA': 'SJEDISTE_PREBIVALISTE_PRIMATELJA',
        'NAČIN OBJAVE': 'NACIN_OBJAVE',
        'VRSTA RASHODA / IZDATKA': 'VRSTA_RASHODA_IZDATKA',
        'Unnamed: 6': 'OPIS_RASHODA',
        'Unnamed: 5': 'OPIS_RASHODA'
    }
    
    # Rename columns that exist in the dataframe
    df = df.rename(columns={k: v for k, v in rename_dict.items() if k in df.columns})
    
    # Remove 'Ukupno' rows if 'NAZIV_PRIMATELJA' column exists
    if 'NAZIV_PRIMATELJA' in df.columns:
        df = df[df['NAZIV_PRIMATELJA'] != 'Ukupno']
    
    # Reset the index
    df = df.reset_index(drop=True)
    
    return df

# Define the path to the documents directory
docs_path = 'dokumenti'

# List all files in the directory
files = os.listdir(docs_path)

# Filter for excel files
excel_files = [f for f in files if f.endswith('.xls') or f.endswith('.xlsx')]

# Check if there are any excel files
if not excel_files:
    print("No Excel files found in the 'dokumenti' directory.")
else:
    # Create an empty list to store dataframes
    all_data = []
    
    # Loop through all excel files
    for file in excel_files:
        file_path = os.path.join(docs_path, file)
        try:
            # Read the excel file
            df = pd.read_excel(file_path, header=4)
            
            # Clean the dataframe
            df = clean_dataframe(df)
            
            # Add a column for the source file
            df['IZVOR'] = file
            
            # Append the dataframe to the list
            all_data.append(df)
        except Exception as e:
            print(f"An error occurred while processing {file}: {e}")
            
    # Concatenate all dataframes
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        
        # Print the info of the final dataframe
        print("Info of the final dataframe:")
        final_df.info()
        
        # Print the first 5 rows of the final dataframe
        print("\nFirst 5 rows of the final dataframe:")
        print(final_df.head())

        # Save the final dataframe to a CSV file
        final_df.to_csv('analiza.csv', index=False)
        print("\nFinal dataframe saved to analiza.csv")