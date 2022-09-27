from pathlib import Path
import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui


def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("Filepath not correct")
    return False


def display_excel_file(excel_file_path, sheet_name):
    df = pd.read_excel(excel_file_path, sheet_name)
    filename = Path(excel_file_path).name
    sg.popup_scrolled(df.types, "=" * 50, df, title=filename)


def convert_to_cv(excel_file_path, output_folder, sheet_name, separator, decimal):
    df = pd.read_excel(excel_file_path, sheet_name)
    filename = Path(excel_file_path).stem
    outputfile = Path(output_folder) / f"{filename}.csv"
    df.to_csv(outputfile, sep=separator, decimal=decimal, index=False)
    sg.popup_no_titlebar("Done! :)")


def settings_window(settings):
    # ------- GUI Definition ------ #
    layout = [
        [sg.Text("SETTINGS")],
        [sg.Text("Separator"), sg.Input(settings["CSV"]["separator"], s=1, key="-SEPARATOR-"),
         sg.Text("Decimal"), sg.Combo(settings["CSV"]["decimal"].split("|"),
                                      default_value=settings["CSV"]["decimal_default"], s=1, key="-DECIMAL-"),
         sg.Text("Sheet Name:"), sg.Input(settings["EXCEL"]["sheet_name"], s=20, key="-SHEET_NAME-")],
        [sg.Button("Save Current Settings", s=20)]]

    window = sg.Window("Settings Window", layout, modal=True)
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == "Save Curent Settings":
            # Write to ini file
            settings["CSV"]["separator"] = values["-SEPARATOR-"]
            settings["CSV"]["decimal_default"] = values["-DECIMAL-"]
            settings["EXCEL"]["sheet_name"] = values["-SHEET_NAME-"]

            # Display success message & close window
            sg.popup_no_titlebar("Settings saved!")
            break
    window.close()


def main_window():
    # ------ Menu Definition ------ #
    menu_def = [["Toolbar", ["Command 1", "Command 2", "---", "Command 3", "Command 4"]],
                ["Help", ["Settings", "About", "Exit"]]]

    # ------ GUI Definition ------ #
    layout = [[sg.MenubarCustom(menu_def, tearoff=False)],
              [sg.Text("Input File:"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xls*"),))],
              [sg.Text("Output Folder:"), sg.Input(key="-OUT-"), sg.FolderBrowse()],
              [sg.Exit(), sg.Button("Settings"), sg.Button("Display Excel File"), sg.Button("Convert To CSV")]]

    window_title = settings["GUI"]["title"]
    window = sg.Window("Excel 2 CSV Converter", layout, use_custom_titlebar=True)

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break
        if event == "About":
            window.disappear()
            sg.popup(window_title, "Version 1.0", "Convert Excel files to CSV", grab_anywhere=True)
            window.reappear()
        if event in ("Command 1","Command 2","Command 3","Command 4"):
            sg.popup("Not yet implemented")
        if event == "Display Excel File":
            if is_valid_path(values["-IN-"]):
                display_excel_file(values["-IN-"], settings["EXCEL"]["sheet_name"])
        if event == "Settings":
            settings_window(settings)
        if event == "Convert TO CSV":
            if is_valid_path(values["-IN-"]) and is_valid_path(values["-OUT-"]):
                convert_to_cv(excel_file_path=values["-IN-"],
                              output_folder=values["-OUT-"],
                              sheet_name=settings["EXCEL"]["sheet_name"],
                              separator=settings["CSV"]["separator"],
                              decimal=settings["CSV"]["decimal"],
                              )

    window.close()


if __name__ == "__main__":
    SETTINGS_PATH = Path.cwd()
    # create the settings object and use ini format
    settings = sg.UserSettings(
        path=SETTINGS_PATH, filename="config.ini", use_config_file=True, convert_bools_and_none=True
    )
    theme = settings["GUI"]["theme"]
    font_family = settings["GUI"]["font_family"]
    font_size = settings["GUI"]["font_size"]
    sg.theme(theme)
    sg.set_options(font=(font_family, font_size))
    main_window()
