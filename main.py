import openpyxl
from openpyxl.styles import Font


def create_empty_file(file: str, name: str, time: str) -> None:
    """
    Creates empty .xlsx file with predefined column names and sets
    the font to Arial Narrow size 10,5pt.

    :param file: file name
    :param name: name for column B, where task names will be stored
    :param time: name for column C, where task times will be stored
    """

    global arial_narrow_font
    global arial_narrow_bold_font
    workbook = openpyxl.Workbook()
    empty_sheet = workbook.active
    empty_sheet.column_dimensions['B'].width = 35
    empty_sheet.column_dimensions['B'].font = arial_narrow_font
    empty_sheet['B2'] = name
    empty_sheet['B2'].font = arial_narrow_bold_font
    empty_sheet.column_dimensions['C'].width = 10
    empty_sheet.column_dimensions['C'].font = arial_narrow_font
    empty_sheet['C2'] = time
    empty_sheet['C2'].font = arial_narrow_bold_font
    workbook.save(file)
    workbook.close()


def find_cell(ws: openpyxl.worksheet,
              column_range: int, row_range: int,
              name: str) -> openpyxl.cell:
    """Find cell in Excel sheet with `name` value.

    :param ws: sheet to be searched through
    :param column_range: maximum non-empty column
    :param row_range: maximum non-empty row
    :param name: value to search for
    :return: cell with value `name`. `None` if it's not found.
    """
    for column in range(1, column_range + 1):
        for row in range(1, row_range + 1):
            temp_cell = ws.cell(row=row, column=column)
            if temp_cell.value == name:
                return temp_cell
    return None


def write_data(task: str, time: int) -> None:
    """

    :param task:
    :param time:
    :return:
    """
    global active_sheet
    global arial_narrow_bold_font
    global arial_narrow_font

    existing_cell = find_cell(active_sheet, max_c, max_r, new_task)

    if existing_cell is None:
        new_task_cell = f"{task_column}{max_r + 1}"
        new_time_cell = f"{time_column}{max_r + 1}"
        active_sheet[new_task_cell] = task
        active_sheet[new_task_cell].font = arial_narrow_font
        active_sheet[new_time_cell] = time
        active_sheet[new_time_cell].font = arial_narrow_font
    else:
        task_row = existing_cell.row
        task_cell = f"{time_column}{task_row}"
        existing_value = active_sheet[task_cell].value
        time += existing_value
        active_sheet[task_cell] = time
        active_sheet[task_cell].font = arial_narrow_font


file_name = "time_example.xlsx"
task_name = "task name"
task_time = "task time"
arial_narrow_bold_font = Font(name="Arial Narrow", size=10.5, bold=True)
arial_narrow_font = Font(name="Arial Narrow", size=10.5)

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


task_column = find_cell(active_sheet, max_c, max_r, task_name).column_letter
time_column = find_cell(active_sheet, max_c, max_r, task_time).column_letter

new = '1'

while new == '1':

    max_r = active_sheet.max_row
    max_c = active_sheet.max_column

    new_task = input("Enter task name: ")
    new_time = int(input("Enter task time: "))

    write_data(new_task, new_time)

    wb.save(file_name)

    new = input("Enter 1 for new task or anything else to exit: ")

wb.close()
