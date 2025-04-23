# Shared Folder File Monitor

A Python program to track and monitor new files added to shared folders by colleagues. This tool helps maintain awareness of file additions and modifications within specified date ranges.

## Purpose

This tool was developed to:
- Monitor shared folders for new file additions
- Track file modifications within specific time periods
- Generate Excel reports of file changes
- Maintain transparency in collaborative work environments

## Features

- Scans specified directories and subdirectories
- Filters files based on modification dates
- Generates detailed Excel reports including:
  - Folder name (oldest level)
  - File name
  - Full file path
  - Modified date
  - Entered date
- Handles large directory structures efficiently
- Provides clear error handling for inaccessible files

## Requirements

- Python 3.6 or higher
- Required packages:
  - pandas
  - openpyxl

## Installation

1. Clone this repository
2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Open `file_monitor.py`
2. Modify the following variables in the `main()` function:
   - `root_path`: Path to your shared folder
   - `begin_date`: Start date for monitoring (YYYY, MM, DD)
   - `end_date`: End date for monitoring (YYYY, MM, DD)
   - `output_path`: Where to save the Excel report

3. Run the script:
```bash
python file_monitor.py
```

## Example Report

The generated Excel report will include:
- Folder Name: The oldest level folder containing the file
- File Name: Name of the file
- Full Path: Complete path to the file
- Modified Date: Last modification timestamp
- Entered Date: File creation timestamp
