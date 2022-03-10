import gspread
import pandas as pd


gc = gspread.oauth()

sheet = gc.open("CaseStudysTest")
# sheet = gc.open_by_url("https://docs.google.com/spreadsheets/d/19Z64MQyVGiypOu0C0OifjyOurbtiV924/edit#gid=2048276846")

ws = sheet.get_worksheet(0)

everything = ws.get_all_records()

df = pd.DataFrame.from_records(everything)[8:]

print(df.head())