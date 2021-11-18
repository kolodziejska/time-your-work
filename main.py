import openpyxl
from openpyxl.styles import Font
import datetime


def create_empty_file(file: str, task: str, time: str) -> None:
    """
    Creates empty .xlsx file with predefined column names and sets
    the font to Arial Narrow size 10,5pt.

    :param file: file name
    :param task: name for column B, where task names will be stored
    :param time: name for column C, where task times will be stored
    """
    global arial_narrow_font
    global arial_narrow_bold_font

    workbook = openpyxl.Workbook()
    empty_sheet = workbook.active

    # set name, font and width for column B,
    # where task names will be stored
    empty_sheet.column_dimensions['B'].width = 35
    empty_sheet.column_dimensions['B'].font = arial_narrow_font
    empty_sheet['B2'] = task
    empty_sheet['B2'].font = arial_narrow_bold_font

    # set name, font and width for column C,
    # where task times will be stored
    empty_sheet.column_dimensions['C'].width = 10
    empty_sheet.column_dimensions['C'].font = arial_narrow_font
    empty_sheet['C2'] = time
    empty_sheet['C2'].font = arial_narrow_bold_font

    workbook.save(file)
    workbook.close()


def find_cell(ws: openpyxl.worksheet, column_range: int, row_range: int,
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


def write_data(ws: openpyxl.worksheet, task: str, time: int) -> None:
    """
    Add task and its time (in hours) to Excel sheet.

    :param ws: sheet to write to.
        The sheet should have pre-defined column names:
        `task_name` for column storing tasks' names and
        `task_time` for column storing tasks' times
    :param task: name of the task
    :param time: time of the task in seconds

    """
    global arial_narrow_bold_font
    global arial_narrow_font
    global task_name
    global task_time

    # find boundaries of non-empty cells in sheet `ws`
    max_r = ws.max_row
    max_c = ws.max_column

    # find in which columns task names and task times are stored
    task_column = find_cell(ws, max_c, max_r, task_name).column_letter
    time_column = find_cell(ws, max_c, max_r, task_time).column_letter

    # check if task with given name already exists
    existing_cell = find_cell(ws, max_c, max_r, task)

    if existing_cell is None:
        # if it doesn't exist, create it in new row
        new_task_cell = f"{task_column}{max_r + 1}"
        new_time_cell = f"{time_column}{max_r + 1}"
        ws[new_task_cell] = task
        ws[new_task_cell].font = arial_narrow_font
        ws[new_time_cell] = time / 3600  # write time in hours
        ws[new_time_cell].font = arial_narrow_font
        ws[new_time_cell].number_format = '0.00'
    else:
        # if it exists, add new time to existing time
        task_row = existing_cell.row
        time_cell = f"{time_column}{task_row}"
        existing_value = ws[time_cell].value  # existing_value in hours
        time += existing_value * 3600  # time in seconds
        ws[time_cell] = time / 3600  # write time back in hours
        ws[time_cell].font = arial_narrow_font
        ws[time_cell].number_format = '0.00'


def task_timer() -> int:
    """Runs simple timer and returns time in seconds."""
    input("press any key to start timer")
    start_time = datetime.datetime.utcnow()
    input("press any key to stop timer")
    end_time = datetime.datetime.utcnow()

    total_time = (end_time - start_time).total_seconds()
    return round(total_time)


def list_existing_tasks(ws: openpyxl.worksheet) -> None:
    """Prints the names of the existing tasks in sheet `ws`."""

    global task_name
    # find boundaries of non-empty cells in sheet `ws`
    max_r = ws.max_row
    max_c = ws.max_column

    # find cell with value `task_name`
    task_cell = find_cell(ws, max_c, max_r, task_name)

    start_row = task_cell.row + 1
    search_column = task_cell.column

    print("\nExisting tasks:")

    for row in range(start_row, max_r + 1):
        task = ws.cell(row=row, column=search_column).value
        print(f"\t{task}")


if __name__ == '__main__':

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

    next_task = '1'

    while next_task == '1':

        new_time = task_timer()
        list_existing_tasks(active_sheet)
        new_task = input("Enter task name: ")

        write_data(active_sheet, new_task, new_time)

        wb.save(file_name)

        next_task = input("Enter 1 for new task or anything else to exit: ")

    wb.close()
