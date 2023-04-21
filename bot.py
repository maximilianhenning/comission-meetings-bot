import pandas as pd
import requests
import openpyxl
import xlrd
from os import path, makedirs, remove
from glob import glob
from datetime import datetime
import re

dir = path.dirname(__file__)

# Load data
links_df = pd.read_csv(path.join(dir, "links.csv"), sep = ";", encoding = "utf-8")
posted_df = pd.read_csv(path.join(dir, "meetings_posted.csv"), sep = ";", encoding = "utf-8")

# Create dictionary
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

def get_meeting_details(meeting, name):
    if "name" in meeting[1].keys():
        meeting_name = meeting[1]["name"]
        category = "cabinet"
    else:
        meeting_name = "nan"
        category = "commissioner"
    date = meeting[1]["date"]
    year = meeting[1]["year"]
    month = meeting[1]["month"]
    day = meeting[1]["day"]
    met_with = meeting[1]["met_with"]
    subject = meeting[1]["subject"]
    return [name, category, meeting_name, date, year, month, day, met_with, subject]

# Get list of meetings to post
meetings_to_post_list = []
for name in names:
    df = pd.read_excel(path.join(dir, "meetings_dumped", name + ".xlsx"))
    # Split date column
    df[["day", "month", "year"]] = df["Date of meeting"].str.split("/", expand = True).astype(int)
    # Select everything that happened after March 2023
    selected_df = df.loc[df["year"] > 2022]
    selected_df = selected_df.loc[selected_df["month"] > 3]
    # Rename columns
    selected_df.rename(columns = {"Name": "name", "Date of meeting": "date", "Location": "location", "Entity/ies met": "met_with", "Subject(s)": "subject"}, inplace = True)
    selected_df["subject"] = selected_df["subject"].str.strip()
    # Check only against relevant part of posted meetings
    check_df = posted_df.loc[posted_df["name"] == name]
    # Check if meetings are in posted meetings already, if not get their info
    for meeting in selected_df.iterrows():
        # Add if there is no meeting on that date yet
        if meeting[1]["date"] not in check_df["date"]:
            meetings_to_post_list.append(get_meeting_details(meeting, name))
        # Or if there is no meeting on that date with that organisation yet
        elif meeting[1]["met_with"] not in check_df.loc[check_df["date"] == meeting[1]["date"]]["met_with"]:
            meetings_to_post_list.append(get_meeting_details(meeting, name))
# Put everything together
to_post_df = pd.DataFrame(meetings_to_post_list)
to_post_df.rename(columns = {0: "commissioner", 1: "category", 2: "persons", 3: "date", 4: "year", 5: "month", 6: "day", 7:"met_with", 8: "subject"}, inplace = True)

# Check if register file needs to be updated and do it if yes
month = datetime.today().strftime("%Y-%m")
register_file = glob(path.join(dir, "register/*"))
last_update = register_file[0].split("\\")[-1].split(".")[0]
if month > last_update:
    resp = requests.get("https://ec.europa.eu/transparencyregister/public/consultation/statistics.do?action=getLobbyistsExcel&fileType=XLS_NEW")   
    with open(path.join(dir, "register", month + ".xls"), "wb") as file:
        file.write(resp.content)
    remove(register_file)
register_file = glob(path.join(dir, "register/*"))[0]
register_df = pd.read_excel(register_file)

# Get register links for organisations
def find_link(met_with):
    register_link_root = "https://ec.europa.eu/transparencyregister/public/consultation/displaylobbyist.do?id="
    name = re.sub(r"\s?\[.*?\]", "", met_with)
    name_match = register_df.loc[register_df["Name"] == name]
    if len(name_match.index) == 1:
        id = name_match["Identification code"].values
    else:
        acronym = re.sub(r"[^\(\)]*(\([^\(\)]*?\))[^\(\)]*", "", met_with)
        acronym_match = register_df.loc[register_df["Acronym"] == acronym]
        if len(acronym_match.index) == 1:
            id = acronym_match["Identification code"].values
    if "id" in locals():
        link = register_link_root + str(id[0])
        return link
    else:
        return ""
to_post_df["link"] = to_post_df["met_with"].apply(find_link)

# Construct messages
for meeting in to_post_df.iterrows():
    # Get variables
    commissioner = meeting[1]["commissioner"]
    category = meeting[1]["category"]
    persons = meeting[1]["persons"]
    date = meeting[1]["date"]
    met_with = meeting[1]["met_with"]
    subject = meeting[1]["subject"]
    link = meeting[1]["link"]
    # Put everything together
    if category == "cabinet":
        message = "Cabinet members of Commissioner " + str(commissioner)
    else:
        message = "Commissioner " + str(commissioner)
    message += " met on " + str(date) + " with:\n\n" + str(met_with) + " " + str(link) + "\n\nSubject(s):\n\n" + str(subject)
    commissioner_code = commissioner[:3]
    general_tag = commissioner_code + "meetings"
    if category == "cabinet":
        specific_tag = commissioner_code + "cab" + "meetings"
    else:
        specific_tag = commissioner_code + "per" + "meetings"
    # XXXX Add tag for met_with
    message += "\n\n#" + general_tag + " #" + specific_tag
    print(message)
    # Add to list

# Post message
    # Wait one minute


# Add meetings to posted file


