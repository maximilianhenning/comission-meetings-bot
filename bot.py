import pandas as pd
import requests
import openpyxl
from os import path, makedirs

dir = path.dirname(__file__)

# Load data
links_df = pd.read_csv(path.join(dir, "links.csv"), sep = ";", encoding = "utf-8")
posted_df = pd.read_csv(path.join(dir, "meetings_posted.csv"), sep = ";", encoding = "utf-8")

# Download new meetings files
link_dict = {}
names = []
for row in links_df.iterrows():
    names.append(row[1]["name"])
    link_dict[row[1]["name"]] = row[1]["link"]

# Get all meetings files
#for name in names:
#    resp = requests.get(link_dict[name])
#    with open(path.join(dir, "meetings_dumped", name + ".xlsx"), "wb") as file:
#        file.write(resp.content)

# Delete first row
#for name in names:
#    wb = openpyxl.load_workbook(path.join(dir, "meetings_dumped", name + ".xlsx"))
#    sheet = wb['Sheet1']
#    sheet.delete_rows(sheet.min_row, 1)
#    wb.save(path.join(dir, "meetings_dumped", name + ".xlsx"))

# Select everything that happened after March 2023
for name in names:
    df = pd.read_excel(path.join(dir, "meetings_dumped", name + ".xlsx"))
    # Split date column
    df[["day", "month", "year"]] = df["Date of meeting"].str.split("/", expand = True).astype(int)
    selected_df = df.loc[df["year"] > 2022]
    selected_df = selected_df.loc[selected_df["month"] > 3]
    selected_df.rename(columns = {"Name": "name", "Date of meeting": "date", "Location": "location", "Entity/ies met": "met_with", "Subject(s)": "subject"}, inplace = True)
    selected_df["subject"] = selected_df["subject"].str.strip()
    # Check if meetings are in posted meetings already, if not get their info
    meetings_to_add_list = []
    for meeting in selected_df.iterrows():
        if meeting[1]["date"] not in posted_df["date"]:
            if "name" in meeting[1].keys():
                meeting_name = meeting[1]["name"]
                category = "cabinet"
            else:
                meeting_name = "nan"
                category = "commissioner"
            year = meeting[1]["year"]
            month = meeting[1]["month"]
            day = meeting[1]["day"]
            met_with = meeting[1]["met_with"]
            subject = meeting[1]["subject"]
            meetings_to_add_list.append([name, category, meeting_name, year, month, day, met_with, subject])
        elif meeting[1]["met_with"] not in posted_df.loc[posted_df["date"] == meeting[1]["date"]]["met_with"]:
            if "name" in meeting[1].keys():
                meeting_name = meeting[1]["name"]
                category = "cabinet"
            else:
                meeting_name = "nan"
                category = "commissioner"
            year = meeting[1]["year"]
            month = meeting[1]["month"]
            day = meeting[1]["day"]
            met_with = meeting[1]["met_with"]
            subject = meeting[1]["subject"]
            meetings_to_add_list.append([name, category, meeting_name, year, month, day, met_with, subject])
    # Put everything together
    print(meetings_to_add_list)

# Check if register file needs to be updated and do it if yes
#register	register	https://ec.europa.eu/transparencyregister/public/consultation/statistics.do?action=getLobbyistsExcel&fileType=XLS_NEW
#register_entry	register	https://ec.europa.eu/transparencyregister/public/consultation/displaylobbyist.do?id=

# Prepare for posting

# Post each new meeting, including register links

# Add meetings to posted file


