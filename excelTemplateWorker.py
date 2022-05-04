import dataclasses
import openpyxl as xl
import pandas as pd

DATA_SHEET_NAME = "data"

COLUMN_NAMES = ["Beworbene SKU", "Beworbene ASIN", "Impressionen", "Klicks", "Klickrate (CTR)", "Kosten pro Klick (CPC)", "Ausgaben", "7 Tage, Umsatz gesamt (€)", "Gesamtumsatzkosten für Werbung (ACOS)", "Gesamtrendite von Werbeausgaben (Return on Advertising Spend, ROAS)", "7 Tage, Aufträge gesamt (#)", "7Tage, Einheiten gesamt (#)", "7-Tage-Konversionsrate", "7Tage, Beworbene SKU-Einheiten (#)", "7-Tage, Andere SKU-Einheiten (#)", "7Tage, Beworbene SKU-Umsätze (€)", "7-Tage, Andere SKU-Umsätze (€)", "Gruppierung"]

HEADERS_TO_CHANGE = {
    "Keyword-oderProdukt-Targeting": "Keyword",
    "GesamtumsatzfürWerbung(ACoS)": "ACOS ",
    "Verkäufe": "14 Tage, Umsatz gesamt",
    "14Tage,Einheitengesamt": "Einheiten insgesamt",
    "Anzeigegruppe": "Anzeigegruppenname",
    "CampaignName(Informationalonly)": "Kampagnen-Name",
    "BeworbeneSKU": "SKU",
    "BeworbeneASIN": "ASIN",
    "AdGroupName(Informationalonly)": "Anzeigegruppenname",
    "Portfolioname": "Portfolio Name"
}


class ExcelWorker:
    def __init__(self, fpath: str):
        self.file_path = fpath
        self.row_nr = 2
        self.col_names = {h_name: index+1 for index, h_name in enumerate(COLUMN_NAMES)}
        self.double_headers = HEADERS_TO_CHANGE
        self.unknown_headers = []

    def write_data(self, df: pd.DataFrame) -> None:
        unknown_headers = []

        if df.empty:
            return

        for header in df.keys():
            h: str = header.replace(" ", "")
            if h.lower() in self.double_headers.keys():
                df.rename({header: self.double_headers[h.lower()]})

        wb = self.load_wb()
        sheet = wb[DATA_SHEET_NAME]

        for header in df.keys():
            if header in self.col_names:
                for row_index, entry in enumerate(df[header]):
                    sheet.cell(row=row_index+self.row_nr, column=self.col_names[header]).value = entry
            else:
                unknown_headers.append(header)

        if any(unknown_headers):
            for header in unknown_headers:
                col_index = sheet.max_row + 2
                sheet.cell(row=1, column=col_index).value = "NOTFOUND: " + header
                for index, entry in enumerate(df[header]):
                    sheet.cell(row=index + self.row_nr, column=col_index).value = entry

        self.row_nr += len(df.keys()[0])
        wb.save(self.file_path)

    def setup(self):
        wb = self.load_wb()
        if DATA_SHEET_NAME not in wb.sheetnames:
            wb.create_sheet(DATA_SHEET_NAME)
        sheet = wb[DATA_SHEET_NAME]
        sheet.delete_rows(1, sheet.max_row + 1)
        for data in self.col_names:
            sheet.cell(row=1, column=self.col_names[data]).value = data
        wb.save(self.file_path)

    def load_wb(self):
        file = self.file_path.split(".")
        if file[-1] == "xlsm":
            return xl.load_workbook(self.file_path, keep_vba=True, read_only=False)
        else:
            return xl.load_workbook(self.file_path)


# old one with spaces
# HEADERS_TO_CHANGE = {
#     "Keyword- oder Produkt-Targeting": "Keyword",
#     "Gesamtumsatz für Werbung (ACoS)": "ACOS ",
#     "Verkäufe": "14 Tage, Umsatz gesamt",
#     "14 Tage, Einheiten gesamt": "Einheiten insgesamt",
#     "Anzeigegruppe": "Anzeigegruppenname",
#     "Campaign Name (Informational only)": "Kampagnen-Name",
#     "Beworbene SKU": "SKU",
#     "Beworbene ASIN": "ASIN",
#     "Ad Group Name (Informational only)": "Anzeigegruppenname",
#     "Portfolioname": "Portfolio Name"
# }
