from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
import time
from datetime import datetime
import json
from win32com.client import Dispatch

chromeOptions = Options()
chromeOptions.add_argument("--headless")

print("Loading webdriver")
PATH = "C:/Program Files (x86)/chromedriver.exe"
driver = webdriver.Chrome(PATH, options=chromeOptions)

print("Getting the page")
driver.get("https://covid19india.org")
time.sleep(10)

print("Parsing")
pageHTML = driver.page_source
driver.quit()

soup = BeautifulSoup(pageHTML, "html.parser")
rows = soup.find_all("div", attrs={"class": "row"})
rows = rows[1:int(len(rows) - 1)]

data = {
    "Author": "Parag Jyoti Pal",
    "description": "Date-wise Indian Covid Stats in JSON !!!!!",
    "_date": datetime.today().strftime('%Y-%m-%d'),
    "data": []
}

for row in rows:
    all_totals = row.find_all("div", attrs={"class": "total"})
    row_data = {
        "stateName": row.find("div", attrs={"class": "state-name"}).text,
        "confirmed": int(all_totals[0]["title"]),
        "confirmedIncrease": 0 if row.find("div", attrs={"class": "is-confirmed"}).text == '' else int(
            row.find("div", attrs={"class": "is-confirmed"}).text[1:]),
        "active": int(all_totals[1]["title"]),
        "recovered": int(all_totals[2]["title"]),
        "recoveredIncrease": 0 if row.find("div", attrs={"class": "is-recovered"}).text == '' else int(
            row.find("div", attrs={"class": "is-recovered"}).text[1:]),
        "deceased": int(all_totals[3]["title"]),
        "deceasedIncrease": 0 if row.find("div", attrs={"class": "is-deceased"}).text == '' else int(
            row.find("div", attrs={"class": "is-deceased"}).text[1:]),
        "tested": int(all_totals[4]["title"])
    }
    data["data"].append(row_data)

print("Saving the data")
with open("./data/Indian-Covid-Stats-" + datetime.today().strftime('%Y-%m-%d') + ".json", "w") as json_file:
    json.dump(data, json_file, indent=4)


def speak(text):
    speech = Dispatch("SAPI.SpVoice")
    speech.Speak(text)


if len(data["data"]) != 0:
    speak("Done")
else:
    speak("Not done")
