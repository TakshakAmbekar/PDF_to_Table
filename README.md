Find the executable file at https://drive.google.com/file/d/16C7HuqcTGNKeVRPy-Z5LfLtt3j6Q9nlH/view?usp=sharing

# PDF Bank Statement Table Extractor
A Python-based GUI tool that extracts tabular data from bank statement PDF files and converts it into a formatted Excel spreadsheet.

# Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation and Setup](#installation-and-setup)
- [Usage Instructions](#usage-instructions)
- [Code Structure](#code-structure)
- [Error Handling](#error-handling)

# Overview
The PDF Bank Statement Table Extractor is a GUI application built with Tkinter. It extracts transaction details such as date, description, amount, and balance from a bank statement in PDF format and exports the data into an Excel file with proper formatting and styling.

# Features
PDF File Selection: Browse and select a bank statement PDF.

- Excel File Configuration: Auto-generate or manually select the output Excel file path.
- Custom Date Formatting: Choose and preview different date formats.
- Data Extraction: Use pdfplumber to extract text and regex to parse transaction details.
- Data Cleaning & Conversion: Convert and clean dates, amounts, and balances.
- Excel Export: Save the processed data into an Excel file using openpyxl with number formatting and table styles.
- User Feedback: Interactive messages guide users through errors, successes, and further actions.

# Requirements
1. Python 3.x
2. Tkinter (usually included with Python)
3. pdfplumber: PDF text extraction
4. pandas: Data manipulation
5. openpyxl: Excel file generation and formatting
6. babel: Number formatting (optional)

# Installation and Setup
1. Clone the Repository
  ```bash
  git clone https://github.com/yourusername/repository-name.git
  cd repository-name
  ```
2. Install the Dependencies
  ```bash
  pip install pdfplumber pandas openpyxl babel
  ```
3. Run the Application
  ```bash
  python GUI.py
  ```

# Usage Instructions
1. Launch the Application:
Running GUI.py opens the PDF Bank Statement Table Extractor window.

2. Select PDF File:
Click the Browse… button next to the PDF File field and choose your bank statement PDF.
Example PDF:
<p align="center">
  <img src="https://github.com/user-attachments/assets/b18b52fc-a10a-4985-a279-c311265f9ffe" alt="Your Image Description">
</p>

4. Configure Excel Output:
Use the auto-generated Excel path or click the Browse… button next to Excel Output to specify a different location.

5. Choose Date Format:
Select your desired date format from the drop-down menu. A live example of the selected format is displayed.

6. Extract and Generate Excel:
Click Extract Table and Generate Excel. The application:
    - Extracts text using pdfplumber
    - Parses transaction data with regex
    - Cleans and formats dates, amounts, and balances
    - Exports the data to a styled Excel file

7. Completion:
A success message will confirm the creation of the Excel file. You can then choose to extract data from another PDF or exit the application.
Example output file:
<p align="center">
  <img src="https://github.com/user-attachments/assets/31578e3a-b08c-49d3-a26d-d5dbab09d7f8" alt="Your Image Description">
</p>

# Code Structure
## Imports and Dependencies
- Core Modules: re, os
- Data Processing: pandas
- PDF Extraction: pdfplumber
- Excel Processing: openpyxl, babel.numbers
- GUI Components: tkinter, ttk, filedialog, messagebox

## Main Class: PDFTableExtractorGUI
- Constructor (__init__):
  - Initializes the GUI window and widgets.
  - Sets default variables and regex patterns.

- File Browsing:
  - browse_pdf(): Opens file dialog for PDF selection and auto-sets Excel output.
  - browse_excel(): Opens dialog to specify Excel output file.

- Date Formatting:
  update_date_example(): Displays a live example of the chosen date format.

- Extraction Process (run_extraction):
  - Extracts text from PDF.
  - Parses the bank statement to extract transactions.
  - Cleans and formats the data.
  - Exports the data to an Excel file.

- Parsing Function (parse_bank_statement):
  Uses regex to extract transaction details and builds a pandas DataFrame.

- Excel Export (save_to_excel):
  Saves the DataFrame to an Excel file with proper formatting and table styling.

## Application Lifecycle
The Tkinter main loop is initiated, and upon extraction completion, the user can choose to process another file or exit the application.

# Error Handling
- Missing File Selections:
Displays error messages if a PDF file or Excel output path is not provided.

- Date Conversion Issues:
Warns the user if date conversion fails, proceeding with the original dates.

- General Exceptions:
Catches unexpected errors during extraction or file saving and displays them.

