import openpyxl
from openpyxl.styles import Font
import datetime
from gui import *


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


def write_data(file: str, task: str, time: int, task_label: str,
               time_label: str) -> None:
    """
    Add task and its time (in hours) to Excel sheet.

    :param file: file to write to.
        The sheet should have pre-defined column names:
        `task_name` for column storing tasks' names and
        `task_time` for column storing tasks' times
    :param task: name of the task
    :param time: time of the task in seconds
    :param task_label: label of the column containing task names
    :param time_label: label of the column containing time names

    """
    global arial_narrow_bold_font
    global arial_narrow_font

    if file is None or file == "":
        return

    # try opening Excel file.
    try:
        wb = openpyxl.load_workbook(file)
        print("File opened successfully")
    except FileNotFoundError:
        print(f"File {file} does not exist. Creating one...")
        create_empty_file(file, task_label, time_label)
        wb = openpyxl.load_workbook(file)
        print("File created and opened successfully")

    ws = wb.active

    # find boundaries of non-empty cells in sheet `ws`
    max_r = ws.max_row
    max_c = ws.max_column

    # find in which columns task names and task times are stored
    task_column_letter = find_cell(ws, max_c, max_r, task_label).column_letter
    time_column_letter = find_cell(ws, max_c, max_r, time_label).column_letter

    # check if task with given name already exists
    existing_cell = find_cell(ws, max_c, max_r, task)

    if existing_cell is None:
        # if it doesn't exist, create it in new row
        new_task_cell = f"{task_column_letter}{max_r + 1}"
        new_time_cell = f"{time_column_letter}{max_r + 1}"
        ws[new_task_cell] = task
        ws[new_task_cell].font = arial_narrow_font
        ws[new_time_cell] = time / 3600  # write time in hours
        ws[new_time_cell].font = arial_narrow_font
        ws[new_time_cell].number_format = '0.00'
    else:
        # if it exists, add new time to existing time
        task_row = existing_cell.row
        time_cell = f"{time_column_letter}{task_row}"
        existing_value = ws[time_cell].value  # existing_value in hours
        time += existing_value * 3600  # time in seconds
        ws[time_cell] = time / 3600  # write time back in hours
        ws[time_cell].font = arial_narrow_font
        ws[time_cell].number_format = '0.00'

    wb.save(file)
    wb.close()


def list_existing_tasks(file: str, task_label: str) -> list[str]:
    """Returns the list of the names of the existing tasks in Excel file.

    :param file: file name
    :param task_label: label of the column with task names
    :return: list of the names of the existing tasks.
            Returns empty list if file name is not specified or valid.
    """

    if file is None or file == "":
        return []

    # try opening Excel file.
    try:
        wb = openpyxl.load_workbook(file)
        print("File opened successfully")
    except FileNotFoundError:
        print(f"File {file} does not exist. Try again.")
        return []

    ws = wb.active
    all_tasks = []

    # find boundaries of non-empty cells in sheet `ws`
    max_r = ws.max_row
    max_c = ws.max_column

    # find cell with value `task_label`
    task_cell = find_cell(ws, max_c, max_r, task_label)

    start_row = task_cell.row + 1
    search_column = task_cell.column

    for row in range(start_row, max_r + 1):
        next_task = ws.cell(row=row, column=search_column).value
        all_tasks.append(next_task)

    wb.close()

    return all_tasks


if __name__ == '__main__':

    task_names = []
    start_time = None
    section_opened = False
    task_column_disabled = True
    time_column_disabled = True

    arial_narrow_bold_font = Font(name="Arial Narrow", size=10.5, bold=True)
    arial_narrow_font = Font(name="Arial Narrow", size=10.5)

    window = sg.Window('Time Your Work', layout, no_titlebar=True,
                       grab_anywhere=True, font='Arial 10', size=(520, 320),
                       finalize=True, use_default_focus=False,
                       margins=(15, 15))

    style = sg.ttk.Style()
    style.configure("TCombobox", borderwidth=0, relief='flat')

    while True:  # Event Loop
        event, values = window.read()
        print(event, values)  # for logging

        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        if event == 'start':
            start_time = datetime.datetime.utcnow()

        if event == 'stop':
            if start_time is not None:
                end_time = datetime.datetime.utcnow()
                total_time = round((end_time - start_time).total_seconds())
                print("total_time: ", total_time)  # for logging
                window['-TOTAL TIME-'].update(total_time)

        if event == '-FILENAME-':
            file_name = str(values["-FILENAME-"])
            task_column = str(values['-TASK COLUMN-'])
            print("file_name:", file_name)  # for logging
            print("task_column: ", task_column)  # for logging
            print("existing task names: ", task_names)  # for logging
            task_names = list_existing_tasks(file_name, task_column)
            window['-TASK NAMES-'].update(values=task_names)
            print("existing task names: ", task_names)  # for logging

        if event.startswith('-OPEN SEC'):
            section_opened = not section_opened
            window['-OPEN SEC-'].update(SYMBOL_UP if section_opened else SYMBOL_DOWN)
            window['-SEC-'].update(visible=section_opened)

        if event == '-SET TASK COLUMN-':
            task_column_disabled = not task_column_disabled
            print(task_column_disabled)  # for logging
            window['-SET TASK COLUMN-'].update(text='edit' if task_column_disabled else 'set')
            window['-TASK COLUMN-'].update(disabled=task_column_disabled)

        if event == '-SET TIME COLUMN-':
            time_column_disabled = not time_column_disabled
            print(time_column_disabled)  # for logging
            window['-SET TIME COLUMN-'].update(text='edit' if time_column_disabled else 'set')
            window['-TIME COLUMN-'].update(disabled=time_column_disabled)

        if event == 'Save':
            file_name = str(values['-FILENAME-'])
            new_task = str(values['-TASK NAMES-'])
            new_time = int(window['-TOTAL TIME-'].get())
            task_column = str(values['-TASK COLUMN-'])
            time_column = str(values['-TIME COLUMN-'])
            print("file_name:", file_name)  # for logging
            print("new_task:", new_task)  # for logging
            print("task_column:", task_column)  # for logging
            print("time_column:", time_column)  # for logging
            print("new_time: ", type(new_time), new_time)  # for logging
            write_data(file_name, new_task, new_time, task_column, time_column)

        if event == 'New task':
            # reset timer
            start_time = None
            window['-TOTAL TIME-'].update(0)
            # update existing task names list in Combobox -TASK NAMES-
            file_name = str(values["-FILENAME-"])
            task_column = str(values['-TASK COLUMN-'])
            task_names = list_existing_tasks(file_name, task_column)
            window['-TASK NAMES-'].update(values=task_names)

    window.close()
