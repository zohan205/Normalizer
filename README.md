# Normalizer

Excel Normalizer is a Python application to clean and normalize Excel files by removing null columns, reformatting headers, and more. The application includes a graphical user interface (GUI) for easy file selection and execution.

## Features

- Remove columns filled with "null" values.
- Rename and delete specific columns based on header names.
- Move the 'eventTime' column to the first position.
- Adjust column widths based on content.
- Provide a GUI for file selection and execution.

## Requirements

- Python 3.x
- Required Python packages:
  - `openpyxl`
  - `tkinter` (included with Python)
  - `datetime`
  - `sys`
  - `os`
  - `time`
  - `threading`

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/ExcelNormalizer.git
