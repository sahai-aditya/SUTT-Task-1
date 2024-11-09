import os
import json
import pandas


EXCEL_FILE_PATH = os.path.join(os.getcwd(), "timetable-workbook.xlsx")

SHEETS = ["S1", "S2", "S3", "S4", "S5", "S6"]
skip_cols = [4, 5]
keep_cols = [i for i in range(12) if i not in skip_cols]
workbook_reader = pandas.read_excel(EXCEL_FILE_PATH, sheet_name=SHEETS, skiprows=[0, 2])

SHEET_KEYS = [
    "COM COD",
    "COURSE NO.",
    "COURSE TITLE",
    "CREDIT",
    "SEC",
    "INSTRUCTOR-IN-CHARGE / Instructor",
    "ROOM",
    "DAYS & HOURS",
    "MIDSEM DATE & SESSION",
    "COMPRE DATE & SESSION"
]

for SHEET_NAME in SHEETS:
    sheet_df = workbook_reader[SHEET_NAME]
    sheet_df.rename(columns={
        "COURSE NO.": "code",
        "COURSE TITLE": "title",
        "CREDIT": "l",
        "Unnamed: 4": "p",
        "Unnamed: 5": "u",
        "SEC": "section",
        "INSTRUCTOR-IN-CHARGE / Instructor": "instructor",
        "ROOM": "room",
        "DAYS & HOURS": "slots",
    }, inplace=True)

    print(f"\n\n{SHEET_NAME}\n")
    for row in sheet_df.iterrows():
        print(row, end="\n---------------------------------------------------------------\n")
