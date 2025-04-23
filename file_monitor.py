import os
import pandas as pd
from datetime import datetime
from pathlib import Path

def get_files_in_date_range(root_path, begin_date, end_date):
    """
    Get all files modified between begin_date and end_date in the given directory and its subdirectories.
    
    Args:
        root_path (str): The root directory to search
        begin_date (datetime): The start date
        end_date (datetime): The end date
        
    Returns:
        list: List of dictionaries containing file information
    """
    file_list = []
    
    for root, dirs, files in os.walk(root_path):
        for file in files:
            file_path = os.path.join(root, file)
            try:
                # Get file modification time
                mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                
                # Check if file was modified within the date range
                if begin_date <= mod_time <= end_date:
                    # Get the relative path from root_path
                    relative_path = os.path.relpath(root, root_path)
                    # Split the path and get the first level folder
                    folder_name = relative_path.split(os.sep)[0] if relative_path != '.' else os.path.basename(root_path)
                    
                    file_info = {
                        'Folder Name': folder_name,
                        'File Name': file,
                        'Full Path': file_path,
                        'Modified Date': mod_time,
                        'Entered Date': datetime.fromtimestamp(os.path.getctime(file_path))
                    }
                    file_list.append(file_info)
            except (PermissionError, FileNotFoundError):
                continue
                
    return file_list

def generate_excel_report(file_list, output_path):
    """
    Generate an Excel report from the file list.
    
    Args:
        file_list (list): List of dictionaries containing file information
        output_path (str): Path where the Excel file will be saved
    """
    if not file_list:
        print("No files found in the specified date range.")
        return
        
    df = pd.DataFrame(file_list)
    df['Modified Date'] = df['Modified Date'].dt.strftime('%Y-%m-%d %H:%M:%S')
    df['Entered Date'] = df['Entered Date'].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    # Reorder columns
    df = df[['Folder Name', 'File Name', 'Full Path', 'Modified Date', 'Entered Date']]
    
    # Save to Excel
    df.to_excel(output_path, index=False)
    print(f"Report generated successfully at: {output_path}")

def main():
    # Example usage
    root_path = r"path\to\your\shared\folder"  # Replace with your shared folder path
    begin_date = datetime(2024, 1, 1)  # Replace with your start date
    end_date = datetime(2024, 12, 31)  # Replace with your end date
    output_path = r"path\to\save\report.xlsx"  # Replace with your desired output path

    file_list = get_files_in_date_range(root_path, begin_date, end_date)
    generate_excel_report(file_list, output_path)

if __name__ == "__main__":
    main() 