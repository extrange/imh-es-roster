"""
Extract the schedule from the XLSX roster to a JSON file
"""
import json
from pathlib import Path
from typing import Tuple, Union

from dotenv import dotenv_values
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from rich.console import Console
from rich.table import Table

# The row in the sheet to extract dates from
DATE_ROW = 5
MY_NAME = dotenv_values('.env')['MY_NAME']

console = Console()


def find_my_name(ws: Worksheet) -> Tuple[str, int]:
    """Find the row with my schedule"""
    for row in ws.rows:
        for cell in row:
            if str(cell.value).strip() == MY_NAME:
                return (cell.column_letter, cell.row)


def find_filled_cell(row, first=True) -> Union[str, None]:
    """
    If first==True, find first column in the row which has a value
    Else find the last
    """
    for cell in row if first else reversed(row):
        if cell.value:
            return cell.column_letter


def print_schedule(schedule):
    """
    Pretty-print the output JSON for user verification.
    """
    table = Table(header_style="bold magenta")
    table.add_column("Day")
    table.add_column("Shift")
    for day, shift in schedule:
        table.add_row(str(day), shift)
    console.print(table)


def main(file: Path):
    wb = load_workbook(file)
    ws = wb[wb.sheetnames[0]]  # Get first sheet

    # Detect row to extract schedule from
    schedule_row = find_my_name(ws)[1]

    # Detect start and end column of date range
    date_start_col = find_filled_cell(ws[DATE_ROW])
    date_end_col = find_filled_cell(ws[DATE_ROW], False)

    schedule = []
    for date_cell in ws[f"{date_start_col}{DATE_ROW}":f"{date_end_col}{DATE_ROW}"][0]:
        current_col = date_cell.column_letter
        schedule_cell = ws[f"{current_col}{schedule_row}"]
        # Add only if there is a value in my schedule
        if schedule_cell.value:
            schedule.append((date_cell.value, schedule_cell.value))

    print_schedule(schedule)

    output_path = Path(file.stem + ".json")
    json.dump(schedule, output_path.open("w", encoding="utf8"), indent=2)
    console.print(f"Wrote to [cyan]{output_path}[/cyan]")


if __name__ == "__main__":
    main(Path("roster.xlsx"))
