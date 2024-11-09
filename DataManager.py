import os
import json
import pandas


BASE_DIR = os.getcwd()
EXCEL_FILE_PATH = os.path.join(BASE_DIR, "timetable-workbook.xlsx")
JSON_FILE_PATH = os.path.join(BASE_DIR, "output.json")

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

    # renaming columns for ease of development
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

    # df_is_notna made to check if a cell is empty
    df_is_notna = sheet_df.notnull()
    course_data = {}
    section = None

    for index, row in sheet_df.iterrows():
        # entering this if means the code found the line which contains course details
        if df_is_notna.loc[index]["code"] and not course_data.get("course_code"):
            course_data["course_code"] = row["code"]
            course_data["course_title"] = row["title"]
            course_data["credits"] = {
                "lecture": row["l"],
                "practical": row["p"],
                "units": row["u"],
            }

            course_data["sections"] = []

        # entering this if means the code found the row which contains section details
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
                # this approach is used to overcome the problem in thermo slots
                timing = []
                slots_split = row["slots"].split()
                slot_num = None

                while slots_split != []:
                    item = slots_split.pop()

                    if item.isdigit():
                        slot_num = int(item)
                        start_time = 7 + slot_num
                        end_time = start_time + 1

                    else:
                        timing.append({
                            "day": item,
                            "timing": [start_time, end_time]
                        })

                # order of slots gets reversed because items are obtained by popping
                # [::-1] done so that timings are from starting of the week to the end
                section_data["timing"] = timing[::-1]

            elif row["section"][0] == "P":
                day, slot, _ = row["slots"].split()
                section_data["timing"].append({
                    "day": day,
                    "timing": [7 + int(slot), 9 + int(slot)]
                })

        # entering this if means the code found the line which contains only the instructor and no new course/section
        elif not df_is_notna.loc[index]["section"] and row["instructor"] not in section_data["instructors"]:
            section_data["instructors"].append(row["instructor"])

    course_data["sections"].append(section_data)
    data.append(course_data)

# dumping the data into the file
with open(JSON_FILE_PATH, "w") as output_f:
    json.dump(data, output_f, indent=4)