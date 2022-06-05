import pandas as pd
from excelTemplateWorker import ExcelWorker
import PySimpleGUI as sg
from time import time
import os


def workflow(save_file, paths):
    excel_worker = ExcelWorker(save_file)
    excel_worker.setup()
    for file in paths:
        wb = pd.ExcelFile(file)
        for sheet in wb.sheet_names:
            sheet_df = pd.read_excel(file, engine="openpyxl", sheet_name=sheet)
            excel_worker.write_data(sheet_df)
    return excel_worker.unknown_headers


def search_folder(path):
    files = os.listdir(path)
    return [os.path.join(path, f) for f in files if f[-5:] == ".xlsx"]


def start():
    folder = sg.popup_get_folder("Wähle den Ordner in dem die Excel Dateien zum zusammfügen liegen\nAlle Dateien mit der Endung xlsx werden verarbeitet", title="ReportAutomationTemplate")
    save_file_path = sg.popup_get_file("Wähle die Datei zum Speichern aus\nBitte schliesse alle Excel Files die du jetzt und vorher ausgewählt hast\nWenn du auf OK drückst startet das Programm", title="ReportAutomationTemplate")
    with_template = sg.Window("Option", [[sg.T("Alles zusammenführen oder per Template erstellen")], [sg.B("Freestyle"), sg.B("Template")]]).read(close=True)
    if folder != "" and folder is not None and save_file_path != "" and save_file_path is not None:
        files = search_folder(folder)
        start_time = time()
        res = workflow(save_file_path, files)

        sg.PopupOK("Folgende Überschriften wurden nicht gefunden " + res[:-2],
                   title=f"Fertig Dauer: {time() - start_time}")


if __name__ == "__main__":
    start()
