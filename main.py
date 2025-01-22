from excel_handler import ExcelHandler

def compare_excels(file1_path, sheet1_name, file2_path, sheet2_name):
    # Initialize readers for both Excel files
    excel1 = ExcelHandler(file1_path, sheet1_name)
    excel2 = ExcelHandler(file2_path, sheet2_name)

    # Open both workbooks
    excel1.open_workbook()
    excel2.open_workbook()

    try:
        # Get the number of rows for each file (assuming column 1 has all data)
        rows1 = len(excel1.get_column_values(1))
        rows2 = len(excel2.get_column_values(1))

        # Loop through all rows in the first file
        for n in range(1, rows1 + 1):
            val1_col2 = str(excel1.get_cell_value(n, 2))
            val1_col3 = str(excel1.get_cell_value(n, 3))
            val1_col5 = str(excel1.get_cell_value(n, 5))

            # Loop through all rows in the second file
            for m in range(1, rows2 + 1):
                val2_col2 = str(excel2.get_cell_value(m, 2))
                val2_col3 = str(excel2.get_cell_value(m, 3))
                val2_col5 = str(excel2.get_cell_value(m, 5))

                # Check if the values match
                if val1_col2 == val2_col2 and val1_col3 == val2_col3 and val1_col5 == val2_col5:
                    # Write 'Test' in cell (10, 10) in the second file
                    excel2.set_cell_value(10, 10, 'Test')
                    break

    finally:
        # Close both workbooks
        excel1.close_workbook()
        excel2.close_workbook()

if __name__ == "__main__":
    # File paths and sheet names
    file1_path = "file1.xlsx"
    sheet1_name = "Sheet1"
    file2_path = "file2.xlsx"
    sheet2_name = "Sheet1"

    compare_excels(file1_path, sheet1_name, file2_path, sheet2_name)
