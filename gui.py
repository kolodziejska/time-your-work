import PySimpleGUI as sg

SYMBOL_UP = '▲'
SYMBOL_DOWN = '▼'


def collapse(layout_c: list[list], key: str) -> sg.Column:
    """
    Helper function that creates a Column that can be later made hidden.
    Function copied from PySimpleGUI Cookbook found at url:
    https://pysimplegui.readthedocs.io/en/latest/cookbook/#recipe-collapsible-sections-visible-invisible-elements
    :param layout_c: The layout for the section
    :param key: Key used to make this section visible / invisible
    :return: A pinned column that can be places directly into the layout.
    """
    return sg.pin(sg.Column(layout_c, key=key, visible=False))


my_new_theme = {'BACKGROUND': '#f2f2f2',
                'TEXT': '#333333',
                'INPUT': '#ffffff',
                'TEXT_INPUT': '#333333',
                'SCROLL': '#ffffff',
                'BUTTON': ('#333333', '#ffffff'),
                'PROGRESS': ('#f2f2f2', '#ffffff'),
                'BORDER': 0,
                'SLIDER_DEPTH': 0,
                'PROGRESS_DEPTH': 0}

sg.theme_add_new("my_theme", my_new_theme)
sg.theme('my_theme')

collapsing_section = [[sg.Text('Column labels currently in use:')],
                      [sg.Input('task name', key='-TASK COLUMN-', size=35,
                                disabled=True, use_readonly_for_disable=True),
                       sg.Button('edit', key='-SET TASK COLUMN-', size=8)],
                      [sg.Input('task time', key='-TIME COLUMN-', size=35,
                                disabled=True, use_readonly_for_disable=True),
                       sg.Button('edit', key='-SET TIME COLUMN-', size=8)]]

column_1 = [[sg.Image(filename='icon.png', size=(60, 60))],
            [sg.Button('start', font='Arial 16 bold', expand_x=True)],
            [sg.Button('stop', font='Arial 16 bold', expand_x=True)],
            [sg.Sizer(v_pixels=10)],
            [sg.Text('total time', font="Arial 8")],
            [sg.Text('0', key='-TOTAL TIME-', font="Arial 18 bold"),
             sg.Text('s', font="Arial 18 bold")],
            [sg.Sizer(h_pixels=140, v_pixels=10)],
            ]

column_2 = [[sg.Text('File:'),
             sg.Input(size=30, key='-FILENAME-', enable_events=True,
                      readonly=True, disabled_readonly_background_color='#f2f2f2'),
             sg.FileBrowse(size=(8, 1), target='-FILENAME-',
                           file_types=[('ALL Files', '.xlsx'),
                                       ('ALL Files', '.xls')])],
            [sg.Text(SYMBOL_DOWN, enable_events=True, key='-OPEN SEC-'),
             sg.Text('Advanced options', enable_events=True, key='-OPEN SEC TEXT-')],
            [collapse(collapsing_section, '-SEC-')],
            [sg.Sizer(v_pixels=8)],
            [sg.Text('Enter task name or choose existing one from the file:'),
             sg.Sizer(h_pixels=20)],
            [sg.Combo(values=[], enable_events=True, size=(40, 10),
                      key='-TASK NAMES-', font='Arial 10 italic', expand_x=True)],
            [sg.Sizer(h_pixels=254), sg.Button('Save', size=8)],
            [sg.Sizer(h_pixels=380)],
            ]

footer = [[sg.Button('New task', size=8), sg.Button('Exit', size=8)],
          [sg.Sizer(h_pixels=520)]]

layout = [[sg.Frame('', column_1, vertical_alignment='top',
                    element_justification='center', expand_x=True,
                    pad=0, relief=sg.RELIEF_SOLID),
           sg.Column(column_2, vertical_alignment='top',
                     element_justification='left')],
          [sg.Sizer(v_pixels=15)],
          [sg.Column(footer, element_justification='right')],
          ]
