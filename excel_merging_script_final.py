
import pandas as pd
import os

def merge_excel_files(tmt_intensity_file, pxd_file):
    # Load the Excel files
    tmt_intensity_df = pd.read_excel(tmt_intensity_file)
    pxd_df = pd.read_excel(pxd_file)

    # Selecting the required columns from PXD037393_ionfinder
    required_columns_pxd = ['scan', 'protein_ID', 'parent_protein', 'protein_description', 'full_sequence', 
                            'sequence', 'formula', 'parent_mz', 'is_modified', 'modified_residue', 'charge', 'contains_Cit']

    # Filtering out only the required columns from PXD037393_ionfinder
    pxd_filtered_df = pxd_df[required_columns_pxd]

    # Merging the two dataframes on 'scan number'
    merged_df = pd.merge(tmt_intensity_df, pxd_filtered_df, on='scan', how='inner')

    return merged_df

def main():
    # Paths to the input Excel files
    tmt_intensity_file = input("Enter the path to the TMT intensity file: ")
    pxd_file = input("Path to Ionfinder Output file: ")  # Path to ionfinder output file

    # Perform the merge
    merged_df = merge_excel_files(tmt_intensity_file, pxd_file)

    # Removing "Normalized TMT-" from the "Highest Intensity TMT" column
    if "Highest Intensity TMT" in merged_df.columns:
        merged_df["Highest Intensity TMT"] = merged_df["Highest Intensity TMT"].str.replace("Normalized TMT-", "", regex=False)
    
    # Save the merged dataframe in the same directory as the output file
    output_file = os.path.join(os.path.dirname(tmt_intensity_file), 'merged_output.xlsx')
    merged_df.to_excel(output_file, index=False)
    print(f"Merged file saved as: {output_file}")
        
if __name__ == "__main__":
    main()
