import xlwings as xw

class ExcelHandler:
    def __init__(self, file_path, sheet_name):
        """
        Initialize the ExcelHandler with the file path and sheet name.

        Parameters:
            file_path (str): The path to the Excel file.
            sheet_name (str): The name of the sheet to access.
        """
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.wb = None
        self.sheet = None
        self.open_workbook()

    def open_workbook(self):
        """
        Open the workbook and select the specified sheet.
        """
        try:
            self.wb = xw.Book(self.file_path)
            self.sheet = self.wb.sheets[self.sheet_name]
        except Exception as e:
            print(f"An error occurred while opening the workbook: {e}")

    def close_workbook(self):
        """
        Save changes to the workbook and close it.
        """
        try:
            if self.wb:
                self.wb.save()
                self.wb.close()
        except Exception as e:
            print(f"An error occurred while closing the workbook: {e}")

    def get_cell_value(self, m, n):
        """
        Get the value of the cell at the mth row and nth column.

        Parameters:
            m (int): The row number (1-based index).
            n (int): The column number (1-based index).

        Returns:
            The value of the cell at (m, n) or None if not found.
        """
        try:
            # Get the value of the specified cell
            cell_value = self.sheet.cells(m, n).value

            return cell_value

        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def get_range_values(self, start_row, start_col, end_row, end_col):
        """
        Get values from a range of cells.

        Parameters:
            start_row (int): Starting row number (1-based index).
            start_col (int): Starting column number (1-based index).
            end_row (int): Ending row number (1-based index).
            end_col (int): Ending column number (1-based index).

        Returns:
            list: A 2D list of values from the specified range.
        """
        try:
            range_values = self.sheet.range((start_row, start_col), (end_row, end_col)).value
            return range_values
        except Exception as e:
            print(f"An error occurred: {e}")
            return None
        
    def get_row_values(self, row, max_col=1000):
        """
        Get all values in the specified row.

        Parameters:
            row (int): Row number (1-based index).

        Returns:
            list: A list of values in the row.
        """
        try:
            row_values = self.sheet.range(f"1{row}:{max_col}{row}").value
            return row_values
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def get_column_values(self, col, max_row=100):
        """
        Get all values in the specified column.

        Parameters:
            col (int): Column number (1-based index).

        Returns:
            list: A list of values in the column.
        """
        try:
            clo_letter = chr(64 + col)
            return self.sheet.range(f"{clo_letter}1:{clo_letter}{max_row}").value
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def find_string_position(self, target):
        """
        Find the position of a speciffic string in the sheet.

        Parameters:
            target (str): The string to search for.

        Returns:
            list: A list of tuples (row, colum) where the string is found.
        """
        positions = []
        for row in range(1, 50):
            print(f"{row} row Searching..")
            for col in range(1, 650):
                if self.sheet.range((row,col)).value == target:
                    print("The Row is: "+str(row)+" and the column is "+str(col))
                    positions.append((row, col))
                    return positions
                
        return positions
        
    def set_cell_value(self, m, n, value):
        """
        Set the value of the cell at the mth row and nth column.

        Parameters:
            m (int): The row number (1-based index).
            n (int): The column number (1-based index).
            value: The value to set in the cell.
        """
        try:
            self.sheet.cells(m, n).value = value
        except Exception as e:
            print(f"An error occurred: {e}")

    def get_sheet_names(self):
        """
        Get a list of all sheet names in the workbook.

        Returns:
            list: A list of sheet names.
        """
        try:
            wb = xw.Book(self.file_path)
            sheet_names = [sheet.name for sheet in wb.sheets]
            wb.close()
            return sheet_names
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def is_cell_empty(self, m, n):
        """
        Check if the cell at the mth row and nth column is empty.

        Parameters:
            m (int): The row number (1-based index).
            n (int): The column number (1-based index).

        Returns:
            bool: True if the cell is empty, False otherwise.
        """
        value = self.get_cell_value(m, n)
        return value is None

# Example usage
if __name__ == "__main__":
    file_path = r"D:\ADMIN_SUNGJUN\CCIC_RANDOM\IBD\XLwings\Excel\base.xlsx"  # Replace with your file path
    sheet_name = "NAVI"       # Replace with your sheet name

    excel_reader = ExcelHandler(file_path, sheet_name)
    excel_reader.open_workbook()

    # # Test get_cell_value
    # m = 3  # Row number
    # n = 2  # Column number
    # value = excel_reader.get_cell_value(m, n)
    # print(f"Value at row {m}, column {n}: {value}")

    # # Test get_row_values
    # row = 3  # Row number
    # row_values = excel_reader.get_row_values(row)
    # print(f"Values in row {row}: {row_values}")

    # # Test get_column_values
    # col = 2  # Column number
    # column_values = excel_reader.get_column_values(col)
    # print(f"Values in column {col}: {column_values}")

    # # Test get_range_values
    # start_row = 1
    # start_col = 1
    # end_row = 3
    # end_col = 3
    # range_values = excel_reader.get_range_values(start_row, start_col, end_row, end_col)
    # print(f"Values in range ({start_row}, {start_col}) to ({end_row}, {end_col}): {range_values}")

    # 자동화 메뉴얼 구분 (2, 65)
    # Level1 (2, 9)
    # Level2 (2, 10)
    # Level3 (2, 11)
    target_string = 'Level3'
    test = excel_reader.find_string_position(target_string)
    print(f"Level3가 적힌 값: {test}")
    

    # # Test setting a cell value
    # new_value = "Test Value"
    # excel_reader.set_cell_value(5, 5, new_value)
    # print(f"Set value at row 5, column 5: {new_value}")

    # # Test is_cell_empty
    # is_empty = excel_reader.is_cell_empty(5, 5)
    # print(f"Is cell at row 5, column 5 empty? {is_empty}")

    # # Test get_sheet_names
    # sheet_names = excel_reader.get_sheet_names()
    # print(f"Sheet names: {sheet_names}")

    # excel_reader.close_workbook()
