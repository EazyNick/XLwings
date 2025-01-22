from excel_handler import ExcelHandler

def compare_excels(file1_path, sheet1_name, file2_path, sheet2_name):
    # Initialize readers for both Excel files
    excel1 = ExcelHandler(file1_path, sheet1_name)
    excel2 = ExcelHandler(file2_path, sheet2_name)

    try:
        # Get the number of rows for each file (assuming column 1 has all data)
        rows1 = len(excel1.get_column_values(1))
        rows2 = len(excel2.get_column_values(1))

        # Loop through all rows in the first file
        for n in range(1, rows1 + 1):
            val1_level1 = str(excel1.get_cell_value(n, 9))
            val1_level2 = str(excel1.get_cell_value(n, 10))
            val1_level3 = str(excel1.get_cell_value(n, 11))
            val1_select = str(excel1.get_cell_value(n, 65))

            # Loop through all rows in the second file
            for m in range(1, rows2 + 1):
                val2_level1 = str(excel2.get_cell_value(m, 9))
                val2_level2 = str(excel2.get_cell_value(m, 10))
                val2_level3 = str(excel2.get_cell_value(m, 11))
                val2_select = str(excel1.get_cell_value(n, 65))

                # Check if the values match
                if val1_level1 == val2_level1 and val1_level2 == val2_level2 and val1_level3 == val2_level3:
                    # Write 'Test' in cell (10, 10) in the second file
                    print(f"({m}, 65), {val1_select}")
                    excel2.set_cell_value(m, 65, val1_select)
                    break

    finally:
        # Close both workbooks
        excel1.close_workbook()
        excel2.close_workbook()

if __name__ == "__main__":
    # File paths and sheet names
    file1_path = r"D:\ADMIN_SUNGJUN\CCIC_RANDOM\IBD\XLwings\Excel\base.xlsx"
    sheet1_name = "NAVI"
    file2_path = r"D:\ADMIN_SUNGJUN\CCIC_RANDOM\IBD\XLwings\Excel\target.xlsx"
    sheet2_name = "NAVI"

    compare_excels(file1_path, sheet1_name, file2_path, sheet2_name)
