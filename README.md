# Email Content Extraction and Cleaning

## Overview
This Python script is designed for extracting and cleaning email content from HTML files. It reads data from a CSV file, processes each HTML file to extract email content and subject, and then performs text cleaning. The cleaned data is then saved into an Excel file.

## Key Features
- Extracts email content and subject from HTML files.
- Cleans the extracted text by removing illegal characters and preserving line breaks and multiple spaces.
- Reads initial data from a CSV file and outputs the processed data to an Excel file.

## How It Works
1. **Data Loading**: The script starts by loading data from a CSV file into a pandas DataFrame.
2. **Email Extraction and Cleaning**:
   - Extracts email content and subjects from HTML files specified in the DataFrame.
   - Cleans the extracted text to remove unwanted characters and formats.
3. **Output Generation**:
   - The script outputs the cleaned data into an Excel file, preserving the original structure.

## Requirements
- Python
- Pandas
- BeautifulSoup
- Openpyxl
- Regex (re module)
