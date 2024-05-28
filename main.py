import pandas as pd
import numpy as np
import os
from openpyxl import load_workbook
import matplotlib.pyplot as plt

def process_excel_file(file_path, output_folder, stddev_folder):
    # Read all sheets into a dictionary of DataFrames
    sheets_dict = pd.read_excel(file_path, sheet_name=None)
    std_devs = {}

    # Iterate through each sheet and process the data
    for sheet_name, df in sheets_dict.items():

        # Replace 0 with NaN, except for the B2 cell
        df.replace(0, np.nan, inplace=True)
        df.iat[0, 1] = 0  # Restore B2 value

        # Remove columns that are all zeros except for the value in the first row
        cols_to_drop = [col for col in df.columns if df[col].iloc[0] != 0 and df[col].iloc[1:].isna().all()]
        df.drop(columns=cols_to_drop, inplace=True)

        # Fill missing values with the mean of the column
        df.fillna(df.mean(), inplace=True)

        # Exclude the first column and calculate standard deviation for each column
        numeric_df = df.iloc[:, 1:].select_dtypes(include=[np.number])
        if not numeric_df.empty:
            std_devs[sheet_name] = numeric_df.std()

        # Update the dictionary with the imputed DataFrame
        sheets_dict[sheet_name] = df

    # Generate the output file path with "CLEANED(MI)" suffix
    base_name = os.path.basename(file_path)
    name, ext = os.path.splitext(base_name)
    output_file_name = f"{name}_CLEANED(MI){ext}"
    output_file_path = os.path.join(output_folder, output_file_name)

    # Save the cleaned data back to an Excel file
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Adjust column widths to fit the content
    wb = load_workbook(output_file_path)
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width
    wb.save(output_file_path)

    # Ensure the stddev folder exists
    if not os.path.exists(stddev_folder):
        os.makedirs(stddev_folder)

    # Generate standard deviation graph
    for sheet_name, std_dev in std_devs.items():
        if not std_dev.empty:
            std_dev.plot(kind='bar', title=f'Standard Deviation for {sheet_name}')
            plt.ylabel('Standard Deviation')
            plt.xlabel('Columns')
            plt.tight_layout()

            # Save the graph
            graph_output_file_name = f"{name}_{sheet_name}_std_dev.png"
            graph_output_path = os.path.join(stddev_folder, graph_output_file_name)
            plt.savefig(graph_output_path)
            plt.close()

def process_all_excel_files(input_folder, output_folder, stddev_folder):
    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Ensure the standard deviations folder exists
    if not os.path.exists(stddev_folder):
        os.makedirs(stddev_folder)

    # Iterate over all .xlsx files in the input folder
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(input_folder, file_name)
            process_excel_file(file_path, output_folder, stddev_folder)

def main():
    input_folder = 'data'
    output_folder = 'cleaned_data'
    stddev_folder = 'standarddeviations'

    # Process all files in the input folder
    process_all_excel_files(input_folder, output_folder, stddev_folder)

if __name__ == '__main__':
    main()
