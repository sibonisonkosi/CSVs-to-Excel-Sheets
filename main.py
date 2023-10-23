import pandas as pd
import os


def main():
    csv_directory = r"C:\Metadata"

    # List all CSV files in the directory
    csv_files = [file for file in os.listdir(csv_directory) if file.endswith('.txt')]

    # Initialize an Excel writer
    output_excel_file = os.path.join(csv_directory, "Data Dictionary.xlsx")
    with pd.ExcelWriter(output_excel_file, engine='xlsxwriter') as writer:
        for idx, file in enumerate(csv_files, start=1):
            file_path = os.path.join(csv_directory, file)
            df = pd.read_csv(file_path, delimiter='\t')
            sheet_name = file.split('.')[0]
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)
        # Add new Sheet
        file_path = os.path.join(csv_directory, 'MeasuresColumns.txt')
        df = pd.read_csv(file_path, delimiter='\t')
        # Move the first column to the end
        columns = df.columns.tolist()
        columns.append(columns.pop(0))
        df = df[columns]
        sheet_name = 'MeasureTableDependencies'
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)

    print(f"Data successfully merged and saved to {output_excel_file}!")


if __name__ == '__main__':
    main()
