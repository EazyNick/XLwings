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
            # Open the workbook and select the sheet
            wb = xw.Book(self.file_path)
            sheet = wb.sheets[self.sheet_name]

            # Get the value of the specified cell
            cell_value = sheet.cells(m, n).value

            # Close the workbook
            wb.close()

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
            wb = xw.Book(self.file_path)
            sheet = wb.sheets[self.sheet_name]
            range_values = sheet.range((start_row, start_col), (end_row, end_col)).value
            wb.close()
            return range_values
        except Exception as e:
            print(f"An error occurred: {e}")
            return None
        
    def get_row_values(self, row):
        """
        Get all values in the specified row.

        Parameters:
            row (int): Row number (1-based index).

        Returns:
            list: A list of values in the row.
        """
        try:
            wb = xw.Book(self.file_path)
            sheet = wb.sheets[self.sheet_name]
            row_values = sheet.range(f"{row}:{row}").value
            wb.close()
            return row_values
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def get_column_values(self, col):
        """
        Get all values in the specified column.

        Parameters:
            col (int): Column number (1-based index).

        Returns:
            list: A list of values in the column.
        """
        try:
            wb = xw.Book(self.file_path)
            sheet = wb.sheets[self.sheet_name]
            col_values = sheet.range(f"{chr(64 + col)}:{chr(64 + col)}").value
            wb.close()
            return col_values
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def set_cell_value(self, m, n, value):
        """
        Set the value of the cell at the mth row and nth column.

        Parameters:
            m (int): The row number (1-based index).
            n (int): The column number (1-based index).
            value: The value to set in the cell.
        """
        try:
            wb = xw.Book(self.file_path)
            sheet = wb.sheets[self.sheet_name]
            sheet.cells(m, n).value = value
            wb.save()
            wb.close()
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
    file_path = "example.xlsx"  # Replace with your file path
    sheet_name = "Sheet1"       # Replace with your sheet name

    excel_reader = ExcelHandler(file_path, sheet_name)
    excel_reader.open_workbook()

    # Test get_cell_value
    m = 3  # Row number
    n = 2  # Column number
    value = excel_reader.get_cell_value(m, n)
    print(f"Value at row {m}, column {n}: {value}")

    # Test get_row_values
    row = 3  # Row number
    row_values = excel_reader.get_row_values(row)
    print(f"Values in row {row}: {row_values}")

    # Test get_column_values
    col = 2  # Column number
    column_values = excel_reader.get_column_values(col)
    print(f"Values in column {col}: {column_values}")

    # Test get_range_values
    start_row = 1
    start_col = 1
    end_row = 3
    end_col = 3
    range_values = excel_reader.get_range_values(start_row, start_col, end_row, end_col)
    print(f"Values in range ({start_row}, {start_col}) to ({end_row}, {end_col}): {range_values}")

    # Test setting a cell value
    new_value = "Test Value"
    excel_reader.set_cell_value(5, 5, new_value)
    print(f"Set value at row 5, column 5: {new_value}")

    # Test is_cell_empty
    is_empty = excel_reader.is_cell_empty(5, 5)
    print(f"Is cell at row 5, column 5 empty? {is_empty}")

    # Test get_sheet_names
    sheet_names = excel_reader.get_sheet_names()
    print(f"Sheet names: {sheet_names}")

    excel_reader.close_workbook()
