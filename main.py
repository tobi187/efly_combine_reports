import subprocess

import pandas as pd
from ex_worker import ExcelWorker
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


def search_folder(path):
    files = os.listdir(path)
    return [os.path.join(path, f) for f in files if f[-5:] == ".xlsx"]


def start():
    folder = sg.popup_get_folder("Wähle den Ordner in dem die Excel Dateien zum zusammfügen liegen\nAlle Dateien mit der Endung xlsx werden verarbeitet", title="ReportAutomation")
    save_file_path = sg.popup_get_file("Wähle die Datei zum Speichern aus\nBitte schliesse alle Excel Files die du jetzt und vorher ausgewählt hast\nWenn du auf OK drückst startet das Programm", title="ReportAutomation")
    if folder == "" and save_file_path == "":
        update = sg.PopupOKCancel("Update Programm")
        if update == "OK":
            working_dir = os.path.dirname(os.path.realpath(__file__))
            subprocess.Popen(["timeout", "5", "&", "git", "pull"], cwd=working_dir, shell=True)
        else:
            return

    if folder != "" and folder is not None and save_file_path != "" and save_file_path is not None:
        files = search_folder(folder)
        start_time = time()
        workflow(save_file_path, files)
        sg.PopupOK(f"Fertig\nDauer: {time() - start_time}")


if __name__ == "__main__":
    start()