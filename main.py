from pathlib import Path
import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui


def display_excel_file(excel_file_path, sheet_name):
    df = pd.read_excel(excel_file_path, sheet_name)
    filename = Path


def convert_to_cv(excel_file_path, output_folder, sheet_name, separator, decimal):
    df = pd.read_excel(excel_file_path, sheet_name)
    filename = Path(excel_file_path).stem
    outputfile = Path(output_folder) / f"{filename}.csv"
    df.to_csv(outputfile, sep=separator, decimal=decimal, index=False)
    sg.popup_no_titlebar("Done! :)")


# ------ GUI Definition ------ #
layout = [[sg.Text("Input File:"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xls*"),))],
          [sg.Text("Output Folder:"), sg.Input(key="-OUT-"), sg.FolderBrowse()],
          [sg.Exit(), sg.Button("Display Excel File"), sg.Button("Convert To CSV")]]

window = sg.Window("Excel 2 CSV Converter", layout)

while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break
    if event == "Display Excel File":
        pass
    if event == "Convert TO CSV":
        convert_to_cv(excel_file_path=values["-IN-"],
                      output_folder=values["-OUT-"],
                      sheet_name="Sheet1",
                      separator="|",
                      decimal=".")

window.close()
