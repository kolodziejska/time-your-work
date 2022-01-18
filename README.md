## Time Your Work
>This simple script with graphical interface was created to automate timing my work on several concurrent projects (for personal reasons and also curiosity).

![gui](https://github.com/kolodziejska/time-your-work/blob/main/gui.PNG)
___
### Features
- open an Excel file you want to store your time in (more on Excel file format restrictions below),
- determine which column should store task names and which task times,
- start and stop the simple timer,
- save your time to the Excel file (with new or pre-existing task name).

#### Excel file restrictions
The script was written with certain format of Excel file in mind:
- both task names and task times should be stored in named columns;
- both names should be distinct;
- no cell should include an Excel formulae (the script works for raw numbers only).

See time_example.xlsx for an example.

### Dependencies
- [openpyxl v3.0.9](https://pypi.org/project/openpyxl/)
- [PySimpleGui v4.55.1](https://pypi.org/project/PySimpleGUI/)
- [pytz v2021.3](https://pypi.org/project/pytz/)
- [tzlocal v4.1](https://pypi.org/project/tzlocal/)

To install dependencies copy and paste the following into the terminal:

```
pip install openpyxl PySimpleGUI pytz tzlocal
```

### Usage
To start the script, execute main.pyw file.
___

### Project status
This project is completed.
