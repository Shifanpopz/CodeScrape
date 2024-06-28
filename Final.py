import requests
from bs4 import BeautifulSoup
import pandas as pd
import threading
from queue import PriorityQueue
import xlsxwriter
import xlwt
from xlwt.Workbook import *
from pandas import ExcelWriter
import time
s = time.time()

sheets = ["I CSE A","I CSE B","I CSE C","I CCE","I CSBS","I","I AIDS A","I AIDS B","I AIML","I IT","I EEE","I MECH","I ECE A","I ECE B","I ECE C"]

df_list = []

for sheet in sheets:

    df=pd.read_excel(r"C:\Users\shifan\Desktop\Project Repo\Scraping\DATA.xlsx",sheet)

    df_list.append([df,sheet])

for df_val in df_list:

    df = df_val[0]

    sheet = df_val[1]

    links = df.Codechef_Profile_Link

    filtered_results = []

    for link in links:
        url = link
        try:
            response = requests.get(url, timeout=10)
            soup = BeautifulSoup(response.content, 'html.parser')
        except Exception:
            soup = None

        if soup is None or url is None:
            filtered_results.append((url, "NA", "NA", "NA", "NA", "NA", "NA"))
            continue

        v = soup.find_all('div', {'class': 'content'})
        if len(v) < 4:
            filtered_results.append((url, "NA", "NA", "NA", "NA", "NA", "NA"))
        else:
            contest_attended = int(soup.find('div', class_='contest-participated-count').b.text.strip())
            contest_rating = 0
            if contest_attended > 0:
                contest_rating = str(soup.find('div', class_='rating-number').text)[0:4]
            star = soup.find('div', class_='rating-star').span.text.strip()
            division = (soup.find('div', class_='rating-header').find_all('div'))[1].text.strip()
            global_rank = soup.find('strong', class_='global-rank').text
            problems_solved = len(str(soup.find_all('div', {'class': 'content'})[2].text).split(',')) - 1
            print(url, "  ", problems_solved, "  ", division, star, "  ", contest_rating, "  ", global_rank, "  ",
                  contest_attended)
            filtered_results.append((url, problems_solved, division, star, contest_rating, global_rank, contest_attended))

    index = 0

    for value in filtered_results:
        df.loc[index,'No_of_problems_solved'] = value[1]
        df.loc[index, 'Division'] = value[2]
        df.loc[index, 'Star_count'] = value[3]
        df.loc[index, 'Contest_rating'] = value[4]
        df.loc[index, 'Global_Rating'] = value[5]
        df.loc[index, 'Contest_Attended'] = value[6]
        index+=1

    df_val[0] = df

with pd.ExcelWriter('multiple.xlsx', engine='xlsxwriter') as writer:
    for df_val in df_list:
        df = df_val[0]
        df.to_excel(writer, sheet_name=df_val[1])

print(time.time() - s)