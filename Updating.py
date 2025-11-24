import pandas as pd
import os
import glob

def clean_comment(comment):
    """Clean comments by removing 'Discrepancies for Match' lines."""
    if pd.isna(comment) or comment == '':
        return comment
    lines = str(comment).split('\n')
    cleaned_lines = [line for line in lines if not line.strip().startswith('Discrepancies for Match')]
    cleaned = '\n'.join(cleaned_lines).strip()
    return cleaned if cleaned else comment

def update_canf_file(input_path_or_dir="/content/Not pre-calculated ETOFs.xlsx",
                     google_drive_path="/content/drive/MyDrive/DAIRB CARRIER",
                     output_filename="Test amount.xlsx"):
    """
    Process ETOF file from input path/directory and update Excel file in Google Drive.

    Args:
        input_path_or_dir: Path to the input ETOF xlsx file or a directory containing it.
        google_drive_path: Path to Google Drive root (default: /content/drive/MyDrive)
        output_filename: Name of the Excel file to update in Google Drive
    """
    try:
        # Mount Google Drive if in Colab
        import sys
        in_colab = 'google.colab' in sys.modules

        if in_colab:
            # Check if drive is already mounted
            if not os.path.exists("/content/drive"):
                print("Mounting Google Drive...")
                from google.colab import drive
                drive.mount('/content/drive')
                print("Google Drive mounted successfully!")
            else:
                print("Google Drive already mounted.")

        input_file_path = None

        if os.path.isfile(input_path_or_dir):
            input_file_path = input_path_or_dir
        elif os.path.isdir(input_path_or_dir):
            # Look for xlsx files within the directory
            xlsx_files = glob.glob(os.path.join(input_path_or_dir, "*.xlsx"))
            if not xlsx_files:
                xlsx_files = glob.glob(os.path.join(input_path_or_dir, "**", "*.xlsx"), recursive=True)

            if xlsx_files:
                input_file_path = xlsx_files[0]
                if len(xlsx_files) > 1:
                    print(f"Warning: Multiple xlsx files found in '{input_path_or_dir}'. Using: {input_file_path}")
                    print(f"All found files: {xlsx_files}")
            else:
                print(f"Error: No xlsx files found in '{input_path_or_dir}' or its subdirectories.")
                return False
        else:
            print(f"Error: Input path '{input_path_or_dir}' is neither a file nor a directory.")
            return False

        if not input_file_path:
            print("Error: Could not determine input ETOF file path.")
            return False

        print(f"Reading input file: {input_file_path}")

        # Read the ETOF file
        df_etofs = pd.read_excel(input_file_path)

        # Prepare data for Google Sheets: Carrier, Cause of CANF, Amount
        if 'Carrier' in df_etofs.columns and 'Comments' in df_etofs.columns:
            # Create cross-product of Carrier and cleaned Comments
            # First, merge Carrier with cleaned comments
            carrier_cause_df = df_etofs[['Carrier', 'Comments']].copy()
            carrier_cause_df['Comments'] = carrier_cause_df['Comments'].apply(clean_comment)

            # Remove rows with empty comments
            carrier_cause_df = carrier_cause_df[carrier_cause_df['Comments'].notna() & (carrier_cause_df['Comments'] != '')]

            # Count occurrences of each Carrier + Cause combination
            google_sheets_data = carrier_cause_df.groupby(['Carrier', 'Comments']).size().reset_index(name='Amount')
            google_sheets_data.columns = ['Carrier', 'Cause of CANF', 'Amount']
            google_sheets_data = google_sheets_data.sort_values(['Carrier', 'Cause of CANF'])

            print(f"\nPrepared {len(google_sheets_data)} Carrier-Cause combinations for Excel file update.")

            # Path to the existing Excel file in Google Drive
            existing_file_path = os.path.join(google_drive_path, output_filename)

            try:
                # Check if file exists
                if os.path.exists(existing_file_path):
                    # Read existing data from the file
                    existing_df = pd.read_excel(existing_file_path)

                    # Ensure the file has the correct columns
                    if all(col in existing_df.columns for col in ['Carrier', 'Cause of CANF', 'Amount']):
                        # Create a dictionary to store existing amounts by (Carrier, Cause) key
                        existing_amounts = {}
                        for _, row in existing_df.iterrows():
                            key = (str(row['Carrier']).strip(), str(row['Cause of CANF']).strip())
                            try:
                                existing_amounts[key] = int(float(row['Amount']))
                            except:
                                existing_amounts[key] = 0

                        # Update amounts: add new amounts to existing ones
                        for idx, row in google_sheets_data.iterrows():
                            key = (str(row['Carrier']).strip(), str(row['Cause of CANF']).strip())
                            if key in existing_amounts:
                                google_sheets_data.at[idx, 'Amount'] = existing_amounts[key] + row['Amount']

                        # Combine existing data (that's not in new data) with new/updated data
                        existing_keys = set(existing_amounts.keys())
                        new_keys = set(zip(google_sheets_data['Carrier'], google_sheets_data['Cause of CANF']))

                        # Add existing rows that aren't in new data
                        rows_to_add = existing_df[~existing_df.apply(
                            lambda r: (str(r['Carrier']).strip(), str(r['Cause of CANF']).strip()) in new_keys, axis=1
                        )]

                        # Combine data
                        if not rows_to_add.empty:
                            google_sheets_data = pd.concat([google_sheets_data, rows_to_add[['Carrier', 'Cause of CANF', 'Amount']]], ignore_index=True)

                        # Sort by Carrier and Cause
                        google_sheets_data = google_sheets_data.sort_values(['Carrier', 'Cause of CANF']).reset_index(drop=True)

                        # Save updated data back to the file
                        google_sheets_data.to_excel(existing_file_path, index=False)
                        print(f"Successfully updated file '{existing_file_path}' with {len(google_sheets_data)} rows!")
                    else:
                        print(f"File '{existing_file_path}' does not have the required columns (Carrier, Cause of CANF, Amount).")
                        return False
                else:
                    # File doesn't exist, create a new one
                    # Ensure the directory exists
                    os.makedirs(os.path.dirname(existing_file_path), exist_ok=True)

                    google_sheets_data = google_sheets_data.sort_values(['Carrier', 'Cause of CANF']).reset_index(drop=True)
                    google_sheets_data.to_excel(existing_file_path, index=False)
                    print(f"Created new file '{existing_file_path}' with {len(google_sheets_data)} rows!")

            except Exception as e:
                print(f"Error updating file: {str(e)}")
                import traceback
                traceback.print_exc()
                return False
        else:
            print("Carrier or Comments column not found. Skipping Excel file update.")
            return False

        return True

    except Exception as e:
        print(f"Error processing file: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

# Example usage:
update_canf_file(
     input_path_or_dir='/content/Not pre-calculated ETOFs.xlsx',
     google_drive_path='/content/drive/MyDrive/DAIRB CARRIER',
     output_filename='Test amount.xlsx'
 )
