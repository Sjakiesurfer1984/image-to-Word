import pandas as pd
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from typing import List, Optional, Any

# A well-structured function for extracting photo numbers from spreadsheets.
def extract_photo_numbers() -> Optional[List[str]]:
    """
    Prompts the user to select an Excel or CSV file and enter a column name.
    It then extracts and normalizes all photo numbers from that column,
    handling cells with multiple numbers, and returns a single list.
    """
    try:
        # Step 1: Use a graphical dialog to let the user select the spreadsheet.
        root: tk.Tk = tk.Tk()
        root.withdraw()
        file_path: str = filedialog.askopenfilename(
            title="Select Excel or CSV file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )

        if not file_path:
            print("No file selected. Exiting.")
            return None
        
        # Step 2: Use input() to ask the user for the column name.
        column_name: str = input("Please enter the name of the column containing the photo numbers: ")
        if not column_name:
            print("No column name entered. Exiting.")
            return None

        # Step 3: Read the spreadsheet into a pandas DataFrame, checking the file extension.
        file_extension = Path(file_path).suffix.lower()
        if file_extension in ['.xlsx', '.xls']:
            df: pd.DataFrame = pd.read_excel(file_path)
        elif file_extension == '.csv':
            df = pd.read_csv(file_path)
        else:
            print("Unsupported file format. Please select an Excel or CSV file.")
            return None

        all_photo_numbers: List[str] = []

        for item in df[column_name].dropna():
            item_str: str = str(item)
            photo_nrs: List[str] = item_str.split()
            for nr in photo_nrs:
                # Final, bulletproof normalization:
                # 1. Remove all non-alphanumeric characters.
                # 2. Convert to lowercase.
                # 3. Add the hyphen back in a consistent location.
                cleaned_nr: str = ''.join(c for c in nr if c.isalnum()).lower()
                if len(cleaned_nr) > 3:
                    cleaned_nr = cleaned_nr[:3] + '-' + cleaned_nr[3:]
                all_photo_numbers.append(cleaned_nr)
        
        return all_photo_numbers
        

    # A robust script anticipates problems and handles them gracefully.
    except FileNotFoundError:
        print(f"Error: The file at '{file_path}' was not found.")
        return None
    except KeyError:
        print(f"Error: The column '{column_name}' does not exist in the file.")
        print("Please check the column name for typos and try again.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

# This block allows the script to be run on its own for testing.
if __name__ == "__main__":
    extracted_list: Optional[List[str]] = extract_photo_numbers()
    if extracted_list:
        print("\nAll extracted photo numbers:")
        print(extracted_list)
        unique_numbers: List[str] = sorted(list(set(extracted_list)))
        print("\nUnique photo numbers:")
        print(unique_numbers)