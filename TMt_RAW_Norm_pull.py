import pandas as pd
import re
import os

def process_excel(file_path):
    # Load the Excel file
    data = pd.read_excel(file_path)
    print("Initial data loaded.")

    # Identifying the start of the actual data in the Excel file
    header_row_idx = None
    for i, row in data.iterrows():
        if row.str.contains('Spectrum Name').any():
            header_row_idx = i
            break

    if header_row_idx is None:
        print("Header row with 'Spectrum Name' not found.")
        return pd.DataFrame()  # Return an empty DataFrame if header row is not found

    print(f"Header row identified at index: {header_row_idx}")

    # Load the actual data
    actual_data = pd.read_excel(file_path, header=header_row_idx)
    print("Actual data loaded.")

    # Renaming columns for easier access
    column_names = actual_data.iloc[0]
    actual_data.columns = column_names
    actual_data = actual_data[1:]  # Remove the first row which is now the header

    # Identifying columns with 'Normalized'
    normalized_columns = [col for col in actual_data.columns if 'Normalized' in str(col)]
    print(f"Normalized columns found: {normalized_columns}")

    if not normalized_columns:
        print("No 'Normalized' columns found.")
        return pd.DataFrame()

    # Convert these columns to numeric
    for col in normalized_columns:
        actual_data[col] = pd.to_numeric(actual_data[col], errors='coerce')

    # Extracting the specific part from 'Spectrum Name' and creating a new column
    actual_data['scan'] = actual_data['Spectrum Name'].apply(lambda x: re.search(r'(\d+)_\d+$', x).group(1) if re.search(r'(\d+)_\d+$', x) else '')

    # Creating the results DataFrame with the required columns
    results = pd.DataFrame()
    results['Spectrum Name'] = actual_data['Spectrum Name']
    results['scan'] = actual_data['scan']
    for col in normalized_columns:
        results[col] = actual_data[col]

    return results

def main():
    # Get user input for file path
    file_path = input("Enter the path to the Excel file: ")

    # Process the file and generate results
    results = process_excel(file_path)

    if results.empty:
        print("No data to save. The resulting DataFrame is empty.")
    else:
        # Save results to a new Excel file in the same directory as the input file
        output_file = os.path.join(os.path.dirname(file_path), os.path.basename(file_path).rsplit('.', 1)[0] + '_processed.xlsx')
        results.to_excel(output_file, index=False)
        print(f"Processed file saved as: {output_file}")

if __name__ == "__main__":
    main()
