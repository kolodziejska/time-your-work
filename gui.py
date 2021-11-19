import PySimpleGUI as sg


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

column_1 = [[sg.Image(filename='icon.png', size=(60, 60))],
            [sg.Button('start', font='Arial 16 bold', expand_x=True)],
            [sg.Button('stop', font='Arial 16 bold', expand_x=True)],
            [sg.Sizer(v_pixels=10)],
            [sg.Text('total time', font="Arial 8")],
            [sg.Text('0', key='-TOTAL TIME-', font="Arial 18 bold"),
             sg.Text('h', font="Arial 18 bold")],
            [sg.Sizer(h_pixels=140, v_pixels=10)],
            ]

column_2 = [[sg.Text('File:'), sg.Input(size=30, key='-FILENAME-'),
             sg.FileBrowse(size=(8, 1),
                           file_types=[('ALL Files', '.xlsx'),
                                       ('ALL Files', '.xls')])],
            [sg.Text('Enter task name or choose existing one from the file:'),
             sg.Sizer(h_pixels=20)],
            [sg.Combo(values=[], enable_events=True, size=(40, 10),
                      key='-TASK NAMES-', font='Arial 10 italic', expand_x=True)],
            [sg.Button('Save', size=8)],
            [sg.Sizer(h_pixels=380)],
            ]

footer = [[sg.Button('New task', size=8), sg.Button('Exit', size=8)],
          [sg.Sizer(h_pixels=520)]]

layout = [[sg.Frame('', column_1, vertical_alignment='top',
                    element_justification='center', expand_x=True,
                    pad=0, relief=sg.RELIEF_SOLID),
           sg.Column(column_2, vertical_alignment='top',
                     element_justification='right')],
          [sg.Sizer(v_pixels=15)],
          [sg.Column(footer, element_justification='right')],
          ]

if __name__ == '__main__':
    window = sg.Window('Time Your Work', layout, no_titlebar=True,
                       grab_anywhere=True, font='Arial 10', size=(520, 320),
                       finalize=True, use_default_focus=False,
                       margins=(15, 15))

    style = sg.ttk.Style()
    style.configure("TCombobox", borderwidth=0, relief='flat')

    while True:  # Event Loop
        event, values = window.read()
        print(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'start':
            # Update the "output" text element to be the value of "input" element
            window['-TOTAL TIME-'].update(1)

    window.close()