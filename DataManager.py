import os
import json
import pandas


EXCEL_FILE_PATH = os.path.join(os.getcwd(), "timetable-workbook.xlsx")

SHEETS = ["S1", "S2", "S3", "S4", "S5", "S6"]
workbook_reader = pandas.read_excel(EXCEL_FILE_PATH, sheet_name=SHEETS, skiprows=[0, 2])

SECTION_TYPES = {
    "L": "lecture",
    "T": "tutorial",
    "P": "practical"
}

data = []

for sheet_name in SHEETS:
    sheet_df = workbook_reader[sheet_name]
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

    df_is_notna = sheet_df.notnull()
    course_data = {}
    section = None

    for index, row in sheet_df.iterrows():
        if df_is_notna.loc[index]["code"] and not course_data.get("course_code"):
            course_data["course_code"] = row["code"]
            course_data["course_title"] = row["title"]
            course_data["credits"] = {
                "lecture": row["l"],
                "practical": row["p"],
                "units": row["u"],
            }

            course_data["sections"] = []
        
        if df_is_notna.loc[index]["section"]:
            if section:
                course_data["sections"].append(section_data)

            section = row["section"]
            section_data = {
                "section_type": SECTION_TYPES[row["section"][0]],
                "section_number": section,
                "instructors": [row["instructor"]],
                "room": int(row["room"]),
                "timing": []
            }

            if row["section"][0] != "P":
                for day in row["slots"].split()[:-1]:
                    start_time = 7 + int(row["slots"].split()[-1])
                    end_time = start_time + 1
                    section_data["timing"].append({
                        "day": day,
                        "slot": [start_time, end_time]
                    })

            elif row["section"][0] == "P":
                day, slot, _ = row["slots"].split()
                section_data["timing"].append({
                    "day": day,
                    "timing": [7 + int(slot), 9 + int(slot)]
                })

        elif not df_is_notna.loc[index]["section"] and row["instructor"] not in section_data["instructors"]:
            section_data["instructors"].append(row["instructor"])
    
    data.append(course_data)


    print(json.dumps(data, indent=4))