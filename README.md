# Normalizer

Excel Normalizer is a Python application designed to clean and standardize Excel files. It removes null columns, renames specific columns, adjusts column widths, and provides a graphical user interface (GUI) for easy file selection and execution.

## Features

- **Null Column Removal**: Remove columns filled with "null" values.
- **Column Renaming and Deletion**: Rename and delete specific columns based on header names.
- **Column Positioning**: Move the 'eventTime' column to the first position.
- **Column Width Adjustment**: Adjust column widths based on content.
- **Graphical User Interface (GUI)**: Provide a user-friendly interface for file selection and execution.

## Requirements

- **Python 3.x**
- **Required Python Packages**:
  - `openpyxl`: For Excel file handling.
  - `tkinter`: For GUI development (included with Python).
  - `datetime`: For timestamp generation.
  - `sys`: For system operations.
  - `os`: For file system operations.
  - `time`: For time-related operations.
  - `threading`: For multi-threading support.

## Installation

1. **Clone the Repository**:
    ```bash
    git clone https://github.com/zohan205/Normalizer.git
    ```

2. **Navigate to the Project Directory**:
    ```bash
    cd Normalizer
    ```

3. **Create and Activate a Virtual Environment**:
    - Windows:
        ```bash
        python -m venv venv
        .\venv\Scripts\activate
        ```
    - macOS/Linux:
        ```bash
        python3 -m venv venv
        source venv/bin/activate
        ```

4. **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

### Running the Application

Run the script:
```bash
python excel_normalizer.py
