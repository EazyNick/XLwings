# XLwings Excel Automation Project

## Overview
This repository contains Python scripts for automating Excel operations using the `xlwings` library. It includes a main script (`main.py`) for comparing data between two Excel files and a utility class (`excel_handler.py`) for handling Excel-specific operations.

## Files

### 1. `main.py`
This script is the entry point for the project. It compares two Excel sheets and updates one file based on matching data from another.

#### Key Features:
- Compares two Excel files based on specific column values.
- Updates target file cells with values from the base file when matches are found.
- Ensures proper workbook handling and cleanup.

#### Usage:
```bash
python main.py
```

### 2. `excel_handler.py`
This file defines the `ExcelHandler` class, which simplifies common Excel operations such as reading, writing, and searching cells.

#### Key Features:
- Open, read, and write Excel files.
- Retrieve row and column data efficiently.
- Search for specific values within a sheet.
- Check if a cell is empty.

## Requirements
- Python 3.8+
- xlwings library

## Installation
1. Clone this repository:
    ```bash
    git clone <repository_url>
    cd <repository_name>
    ```
2. Install the required dependencies:
    ```bash
    pip install xlwings
    ```

## How to Run
1. Ensure the paths to your Excel files are correctly set in `main.py`:
    ```python
    file1_path = r"D:\\path\\to\\base.xlsx"
    sheet1_name = "Sheet1"
    file2_path = r"D:\\path\\to\\target.xlsx"
    sheet2_name = "Sheet2"
    ```
2. Execute the script:
    ```bash
    python main.py
    ```

## Example Workflow
1. The script reads data from a base Excel file (`base.xlsx`) and a target Excel file (`target.xlsx`).
2. It compares specific columns (e.g., Levels 1, 2, 3) to identify matches.
3. When a match is found, a value from the base file is written to the target file in a specified column.

## Contributions
Feel free to open issues or submit pull requests if you would like to contribute to this project.

## License
This project is licensed under the MIT License.

---

### Contact
For questions or feedback, please reach out to the repository maintainer.
