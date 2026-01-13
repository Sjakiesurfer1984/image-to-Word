# --- Module Imports ---
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import shutil
from typing import List, Set

# --- Core Logic Functions ---

def normalize_photo_numbers(photo_numbers: List[str]) -> Set[str]:
    """
    Takes a list of raw photo number strings and standardizes them.

    This function's single responsibility is to clean the input data into a
    consistent format for reliable comparison later.

    Args:
        photo_numbers (List[str]): A list of photo number strings (e.g., '100_0006', '100-7').

    Returns:
        Set[str]: A set of standardized photo numbers (e.g., {'100-0006', '100-0007'}).
                  Using a set provides highly efficient lookups.
    """
    if not photo_numbers:
        print("Warning: The list of photo numbers is empty.")
        return set() # Return an empty set if there's no input.

    # This 'set comprehension' is a concise way to build a set.
    # For each number in the input list, it performs the cleaning and formatting operations.
    # 1. nr.strip().lower().replace('_', '-'): Standardizes separators and case.
    # 2. num.split('-')[-1]: Isolates the numerical part of the photo name.
    # 3. .zfill(4): Pads the number with leading zeros to ensure a consistent length (e.g., '7' -> '0007').
    # 4. '100-' + ... : Reconstructs the name into the final, standard format.
    normalized_set = {'100-' + num.split('-')[-1].zfill(4) for num in photo_numbers}
    
    return normalized_set


def move_files(file_paths: List[Path], output_directory: Path):
    """
    Moves a list of files to a specified output directory.

    This function's single responsibility is the file system operation. It does not
    decide which files to move; it only executes the move command based on the
    list it is given.

    Args:
        file_paths (List[Path]): A list of Path objects representing the files to be moved.
        output_directory (Path): A Path object for the destination folder.
    """
    # --- Input Validation and Setup ---
    if not file_paths:
        print("No files were scheduled for moving. The list is empty.")
        return

    # Create the output directory. The 'exist_ok=True' argument prevents a crash
    # if the folder has already been created in a previous run.
    try:
        output_directory.mkdir(exist_ok=True)
        print(f"Output directory is ready at: {output_directory}")
    except OSError as e:
        print(f"Error: Could not create output directory '{output_directory}'. Reason: {e}")
        return # Stop execution if we cannot create the destination folder.

    # --- File Moving Loop ---
    moved_count = 0
    skipped_count = 0
    for file_path in file_paths:
        # It's good practice to verify the source file exists before trying to move it.
        # The file could have been moved or deleted by another process.
        if not file_path.exists():
            print(f"-> Warning: Source file not found. Skipping: {file_path.name}")
            skipped_count += 1
            continue # 'continue' skips to the next iteration of the loop.

        # The 'try...except' block makes our script robust. If a single file fails to move
        # (e.g., due to a permissions error), the script will report the error and continue
        # with the rest of the files instead of crashing completely.
        try:
            destination = output_directory / file_path.name
            shutil.move(str(file_path), str(destination))
            print(f"-> Moved: {file_path.name}")
            moved_count += 1
        except Exception as e:
            print(f"-> Error: Could not move {file_path.name}. Reason: {e}")
            skipped_count += 1
    
    print(f"\n--- Operation Summary ---")
    print(f"Successfully moved: {moved_count} file(s).")
    if skipped_count > 0:
        print(f"Skipped or failed: {skipped_count} file(s).")
    print("-------------------------")


# --- Main Execution Block (The Controller) ---
if __name__ == "__main__":
    
    # 1. DEFINE the data from the report.
    # This list contains all the photo numbers mentioned in the inspection notes.
    report_photo_list = ['5593-5599', '5603-5614', '561K-5620', '5621-5627', '5628.5635', '5636-39', '5640', '5641', '5642-43', '5644-48', '5649', '5650-51', '5652', '5657']

    # 2. PROCESS the report data by calling our normalization function.
    required_photo_set = normalize_photo_numbers(report_photo_list)
    print(f"Normalized set of required photos: {required_photo_set}")

    # 3. GET the location of the source images from the user.
    root = tk.Tk()
    root.withdraw()
    source_dir_str = filedialog.askdirectory(title="Select the folder containing ALL images")

    # Exit gracefully if the user cancels the dialog.
    if not source_dir_str:
        print("No source folder selected. Exiting.")
    else:
        source_dir = Path(source_dir_str)
        output_dir = source_dir / "not_used_in_report"

        # 4. FIND all image files in the source directory.
        all_images = [p for p in source_dir.glob('*') if p.is_file() and p.suffix.lower() in ['.jpg', '.jpeg', '.png', '.bmp']]
        print(f"\nFound {len(all_images)} total images in the source folder.")

        # 5. DECIDE which files to keep and which to move.
        # This loop iterates through all found images and categorizes them.
        photos_to_move = []
        for image in all_images:
            # Normalize the filename from the file system to match our report's format.
            normalized_stem = image.stem.lower().replace('_', '-')
            
            # The core logic: if the image name is NOT in our required set, it gets moved.
            if normalized_stem not in required_photo_set:
                photos_to_move.append(image)

        print(f"Identified {len(photos_to_move)} images to be moved.")

        # 6. EXECUTE the move operation by calling our file-moving function.
        move_files(photos_to_move, output_dir)
