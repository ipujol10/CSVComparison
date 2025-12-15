"""Compare Excel files to see where they are different"""

import csv
import tkinter as tk
from tkinter import filedialog as fd


def getFiles() -> list[str]:
    """Ask user to select files to compare"""
    root = tk.Tk()
    root.withdraw()

    selected_files: list[str] = []
    for i in range(2):
        file_path = fd.askopenfilename(
            title=f"Select file number {i+1}",
            filetypes=(("CSV files", "*.csv"),),
        )
        selected_files.append(file_path)

    return selected_files


def excelNumberToColumn(col: int) -> str:
    """Convert a 1-based column number to its Excel column name."""
    if col < 1:
        raise ValueError("Column number must be a positive integer")

    result = ""
    while col > 0:
        # Adjust for 1-based indexing (Excel columns start at 1)
        col -= 1
        remainder = col % 26
        # Convert remainder (0-25) to letter A-Z
        result = chr(ord("A") + remainder) + result
        col //= 26
    return result


def compareFiles(file1_name: str, file2_name: str) -> None:
    """Make the actual comparison"""
    with open(file1_name, "r", encoding="cp1252") as file1, open(file2_name, "r", encoding="cp1252") as file2:
        reader1 = csv.reader(file1)
        reader2 = csv.reader(file2)
        for i, row in enumerate(zip(reader1, reader2)):
            for j, (cell1, cell2) in enumerate(zip(*row)):
                if cell1 != cell2:
                    print(f"{excelNumberToColumn(j+1)}{i+1}")


def main() -> None:
    """Compare 2 Excel files (Previously converted to CSV)"""
    files: list[str] = getFiles()
    compareFiles(*files)  # pylint: disable=E1120


if __name__ == "__main__":
    main()
