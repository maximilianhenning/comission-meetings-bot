from os import path, makedirs, remove, environ
environ["OPENBLAS_NUM_THREADS"] = "1"
environ["OMP_NUM_THREADS"] = "1"
environ["MKL_NUM_THREADS"] = "1"
environ["VECLIB_MAXIMUM_THREADS"] = "1"
environ["NUMEXPR_NUM_THREADS"] = "1"
import pandas as pd
import requests
import openpyxl
import xlrd
from glob import glob
from datetime import datetime
import re
from time import sleep
from atproto import Client

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
for name in names:
    resp = requests.get(link_dict[name])
    if not path.exists(path.join(dir, "meetings_dumped")):
        makedirs(path.join(dir, "meetings_dumped"))
    with open(path.join(dir, "meetings_dumped", name + ".xlsx"), "wb") as file:
        file.write(resp.content)

# Delete first row
for name in names:
    try:
        wb = openpyxl.load_workbook(path.join(dir, "meetings_dumped", name + ".xlsx"))
        sheet = wb.active
        sheet.delete_rows(1)
        if not path.exists(path.join(dir, "meetings_wrangled")):
            makedirs(path.join(dir, "meetings_wrangled"))
        wb.save(path.join(dir, "meetings_wrangled", name + ".xlsx"))
    except:
        print("Error on name", name)

# Get list of meetings to post
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

meetings_to_post_list = []
for name in names:
    df = pd.read_excel(path.join(dir, "meetings_wrangled", name + ".xlsx"))
    # Proceed only if there are meetings
    if len(df.index > 0):
        # Split date column
        df["Date of meeting"] = df["Date of meeting"].str.replace("/", ".")
        df[["day", "month", "year"]] = df["Date of meeting"].str.split(".", expand = True).astype(int)
        # Select everything that happened after March 2023
        selected_df = df.loc[df["year"] > 2022]
        selected_df = selected_df.loc[selected_df["month"] > 3]
        # Rename columns
        selected_df.rename(columns = {"Name": "name", "Date of meeting": "date", "Location": "location", "Entity/ies met": "met_with", "Subject(s)": "subject"}, inplace = True)
        selected_df["subject"] = selected_df["subject"].str.strip()
        # Check only against relevant part of posted meetings
        check_df = posted_df.loc[posted_df["commissioner"] == name]
        # Check if meetings are in posted meetings already, if not get their info
        for meeting in selected_df.iterrows():
            # Add if there is no meeting on that date yet
            if meeting[1]["date"] not in check_df["date"].tolist():
                meetings_to_post_list.append(get_meeting_details(meeting, name))
            # Or if there is no meeting on that date with that organisation yet
            elif meeting[1]["met_with"] not in check_df.loc[check_df["date"] == meeting[1]["date"]]["met_with"].tolist():
                meetings_to_post_list.append(get_meeting_details(meeting, name))
# If there are meetings to post, put everything together
if not meetings_to_post_list:
    print("No new meetings to post")
else:
    to_post_df = pd.DataFrame(meetings_to_post_list)
    to_post_df.rename(columns = {0: "commissioner", 1: "category", 2: "persons", 3: "date", 4: "year", 5: "month", 6: "day", 7: "met_with", 8: "subject"}, inplace = True)
    to_post_df.sort_values(by = ["year", "month", "day"], ascending = False, inplace = True)

    # Check if register file needs to be updated and do it if yes
    def read_register_file():
        return glob(path.join(dir, "register/*"))[-1]

    month = datetime.today().strftime("%Y-%m")
    register_file = read_register_file()
    last_update = register_file.split("\\")[-1].split(".")[0]
    if month > last_update:
        # Commission sometimes updates faulty spreadsheets - in those cases ignore exception & try again next time
        try:
            resp = requests.get("https://ec.europa.eu/transparencyregister/public/consultation/statistics.do?action=getLobbyistsExcel&fileType=XLS_NEW")   
            new_register_df = pd.read_excel(resp.content)
            new_register_df.to_csv(path.join(dir, "register", month + ".csv"), sep =  ";", encoding = "utf-8", index = False)
        except Exception:
            pass
    register_file = read_register_file()
    register_df = pd.read_csv(register_file, sep = ";", encoding = "utf-8")

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
    message_list = []
    for meeting in to_post_df.iterrows():
        # Get variables
        category = meeting[1]["category"]
        if category == "cabinet":
            commissioner = meeting[1]["commissioner"].split("_")[0]
        else:
            commissioner = meeting[1]["commissioner"]
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
        # Hashtag
        commissioner_code = commissioner[:3]
        tag = commissioner_code + "meetings"
        message += "\n\n#" + tag
        # XXXX Add tag for met_with
        # Add to list
        message_list.append(message)

    # Set up connection to Mastodon API
    with open(path.join(dir, "mastodon_token.txt"), "r") as file:
        token = file.read()
    url = "https://eupolicy.social/api/v1/statuses"
    auth = {"Authorization": "Bearer " + str(token)}

    # Post messages to Mastodon
    print("\n\nMastodon\n\n")
    for message in message_list:
        print(message)
        params = {"status": message}
        r = requests.post(url, data = params, headers = auth)
        print(r)
        sleep(15)

    # Set up connection to Bluesky API
    with open(path.join(dir, "bsky_token.txt"), "r") as file:
        token = file.read()
    client = Client(base_url = "https://bsky.social")
    client.login("eulobbybot.bsky.social", token)

    # Post messages to Bluesky
    print("\n\nBluesky\n\n")
    for message in message_list:
        post = client.send_post(message)
        print(post)
        sleep(15)

    # Add meetings to posted file
    posted_df = pd.concat([posted_df, to_post_df])
    posted_df.to_csv(path.join(dir, "meetings_posted.csv"), sep = ";", encoding = "utf-8", index = False)