import openpyxl
from openpyxl.styles import Font


def create_empty_file(file: str, name: str, time: str) -> None:
    """
    Creates empty .xlsx file with predefined column names and sets
    the font to Arial Narrow size 10,5pt.

    :param file: file name
    :param name: value for column B, where task names will be stored
    :param time: value for column C, where task times will be stored
    """

    arial_narrow_font = Font(name="Arial Narrow", size=10.5, bold=True)
    workbook = openpyxl.Workbook()
    empty_sheet = workbook.active
    empty_sheet.column_dimensions['B'].width = 35
    empty_sheet.column_dimensions['B'].font = arial_narrow_font
    empty_sheet['B2'] = name
    empty_sheet['B2'].font = arial_narrow_font
    empty_sheet.column_dimensions['C'].width = 10
    empty_sheet.column_dimensions['C'].font = arial_narrow_font
    empty_sheet['C2'] = time
    empty_sheet['C2'].font = arial_narrow_font
    workbook.save(file)


def find_named_column(ws: openpyxl.worksheet.worksheet.Worksheet,
                     column_range: int, row_range: int,
                     name: str) -> str:
    """Find column in Excel sheet with `name` value.

    :param ws: sheet to be searched through
    :param column_range: maximum non-empty column
    :param row_range: maximum non-empty row
    :param name: value to search for
    :return: letter of the column
    """
    for column in range(1, column_range + 1):
        for row in range(1, row_range + 1):
            temp_cell = ws.cell(row=row, column=column)
            if temp_cell.value == name:
                return temp_cell.column_letter
            else:
                return None


file_name = "time_example.xlsx"
task_name = "task name"
task_time = "task time"

try:
    wb = openpyxl.load_workbook(file_name)
    print("File opened successfully")
except FileNotFoundError:
    print(f"File {file_name} does not exist. Creating one...")
    create_empty_file(file_name, task_name, task_time)
    wb = openpyxl.load_workbook(file_name)
    print("File created and opened successfully")

active_sheet = wb.active
max_r = active_sheet.max_row
max_c = active_sheet.max_column

task_column = find_named_column(active_sheet, max_c, max_r, task_name)
time_column = find_named_column(active_sheet, max_c, max_r, task_time)

new_task = input("Enter task name: ")
new_time = int(input("Enter task time: "))

new_task_cell = f"{task_column}{max_r + 1}"
new_time_cell = f"{time_column}{max_r + 1}"

active_sheet[new_task_cell] = new_task
active_sheet[new_time_cell] = new_time

wb.save(file_name)
