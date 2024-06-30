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
from bs4 import BeautifulSoup

dir = path.dirname(__file__)


# Get Commissioner links
def get_commissioner_links(commission_url):
    commissioners = {}
    response = requests.get(commission_url)
    html = response.text
    doc = BeautifulSoup(html, "html.parser")
    link_master_divs = doc.find_all("div", {"class": "ecl-content-item-block__item ecl-u-mb-l"})
    for link_master_div in link_master_divs:
        try:
            link_div = link_master_div.find("div", {"class": "ecl-content-block__title"})
            commissioner = link_div.find("a").contents[0]
            commissioner = commissioner.split(" ")[1]
            link = link_div.find("a").get("href")
            link = "https://commissioners.ec.europa.eu" + link
            commissioners[commissioner] = link
        except:
            pass
    links_df = pd.DataFrame.from_dict(commissioners, orient = "index")
    links_df = links_df.reset_index().rename(columns = {"index": "name", 0: "link"})
    return links_df

def get_meeting_links(link):
    sleep(5)
    response = requests.get(link)
    html = response.text
    doc = BeautifulSoup(html, "html.parser")
    meeting_links = []
    a_tag_list = doc.find_all("a")
    for a_tag in a_tag_list:
        try:
            href = a_tag.get("href")
            if "transparencyinitiative" in href:
                href = href.split("?")[1]
                href = "https://ec.europa.eu/transparencyinitiative/meetings/exportmeetings.do?" + href
                meeting_links.append(href)
        except:
            pass
    meeting_links = meeting_links[:2]
    return pd.Series(meeting_links, dtype = "object")

def get_category_file(name, category, link):
    response = requests.get(link)
    with open(path.join(dir, "meetings_dumped_new", name + "_" + category + ".xlsx"), "wb") as file:
        file.write(response.content)

def get_meeting_files(row):
    name = row["name"]
    get_category_file(name, "commissioner", row["commissioner"])
    get_category_file(name, "cabinet", row["cabinet"])

def delete_first_row(file):
    filename = path.basename(file).split(".")[0]
    wb = openpyxl.load_workbook(file)
    sheet = wb.active
    sheet.delete_rows(1)
    wb.save(path.join(dir, "meetings_wrangled_new", filename + ".xlsx"))

if not path.exists(path.join(dir, "meetings_dumped_new")):
    makedirs(path.join(dir, "meetings_dumped_new"))
if not path.exists(path.join(dir, "meetings_wrangled_new")):
    makedirs(path.join(dir, "meetings_wrangled_new"))

links_df = get_commissioner_links("https://commissioners.ec.europa.eu/index_en")
links_df[["commissioner", "cabinet"]] = links_df["link"].apply(get_meeting_links)
links_df.to_csv(path.join(dir, "links_new.csv"), sep = ";", index = False, encoding = "utf-8")
links_df.apply(lambda row: get_meeting_files(row), axis = 1)
for file in glob(path.join(dir, "meetings_dumped_new")):
    delete_first_row(file)