#!/usr/bin/env python3
"""
Follow a runner live using the GeeXek (赛客) race timing service (https://www.geexek.com/score).

Asks for the competition, race category and ranking list (or takes them from --cmptId / --rankId),
finds the runner by bib and name, then redraws their checkpoint passing times, speed, rank and
predicted times every 10 seconds until stopped with Ctrl+C.

This is largely a proof of concept to use GeeXek's undocumented backend API.
Error handling and runner selection are minimal.

For handling a full list of race results (when available):

Save results json from race page, e.g. https://www.geexek.com/score/singleSearch?Y2lkPTMyNzExJmNoa0ZsZz0mcm9hZFR5cGU9MiZjbXB0TmFtZT0yMDI2JUU3JUFDJUFDJUU0JUI4JTg5JUU1JUIxJThBJUU5JTg3JThEJUU1JUJBJTg2JUU2JUFEJUE2JUU5JTlBJTg2JUU2JThFJUE1JUU5JUJFJTk5JUU4JUI2JThBJUU5JTg3JThFJUU4JUI1JTlC
Convert with:
  df = pd.read_json('results.json')
  df.to_excel('results.xlsx')
"""

# Standard library
import argparse
import json
import sys
from time import sleep, strftime

# Third party
import pandas as pd
import requests

parser = argparse.ArgumentParser()
parser.add_argument('--cmptId', help='Competition ID', default=None, required=False)
parser.add_argument('--rankId', help='Ranking ID', default=None, required=False)
parser.add_argument('--bib', help='Bib number', default=None, required=False)
parser.add_argument('--name', help='Runner name', default=None, required=False)
args = parser.parse_args()

ua = 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:153.0) Gecko/20100101 Firefox/153.0'
TIMEOUT = 15  # seconds to wait for the timing server before giving up on a request


# Find competitions:
# https://www.geexek.com/score
url = 'https://timerbackend.geexek.com/getCmptListForWeixin.do'
start = 0
cmptId = args.cmptId
while (not cmptId):
    response = requests.post(url, timeout=TIMEOUT, headers={'User-Agent': ua}, data={
        'cmptType': -1,
        'start': start,
        'searchContent': ''
    })
    if response.status_code != 200:
        print(f"Error: {response.status_code}")
        sys.exit(1)

    cmptList = pd.DataFrame(response.json()['cmptList'])
    print(cmptList[['cmptId', 'cmptStartTime', 'cmptCity', 'cmptName', 'scoreOpened', 'cmptType', 'rankShow', 'openedStatus']].to_markdown(index=False))

    cmptId = input("Enter cmptId or enter to load more races: ")
    if not cmptId.isdigit():
        cmptId = None
    start += 1

rankId = args.rankId
if not rankId:
    # Find ranking lists for competition:
    url = 'https://timerbackend.geexek.com/raceBoard/scoreBoardInfo.do'
    response = requests.post(url, timeout=TIMEOUT, headers={'User-Agent': ua}, data={'cmptId': cmptId})
    if response.status_code != 200:
        print(f"Error: {response.status_code}")
        sys.exit(1)
    # print(json.dumps(response.json(), indent=4, ensure_ascii=False))
    roadList = pd.DataFrame(response.json()['roadList'])
    print(roadList[['roadName', 'roadTypeLimit', 'roadType', 'isPublish']].to_markdown())
    road = None
    while road is None:
        try:
            road = int(input("Enter category: "))
        except ValueError:
            continue
        if not 0 <= road < len(roadList):
            road = None
    rankList = pd.DataFrame(roadList.iloc[road]['rankList'])
    print(rankList[['rankId', 'listName', 'rankType', 'userGender']].to_markdown(index=False))
    rankId = None
    while (not rankId):
        rankId = input("Enter rankId: ")
        if not rankId.isdigit():
            rankId = None

#List runners in the ranking list:
url = 'https://timerbackend.geexek.com/rankingBoard.do'
pageNo = 1 # unfortunately incrementing the pageNo doesn't work, it always returns the same results, so we can only get the first 200 runners
response = requests.post(url, timeout=TIMEOUT, headers={'User-Agent': ua}, data={'rankId': rankId, 'pageNo': pageNo})
if response.status_code != 200:     
    print(f"Error: {response.status_code}")
    sys.exit(1)
runners = pd.DataFrame(response.json()['list'])
print(runners[['id', 'mark', 'name', 'length', 'state', 'realTimeFormat']].to_markdown())

#Search for a runner
markNo = args.bib
if not markNo:
    markNo = input("Bib number: ")
keyword = args.name
if not keyword:
    keyword = input("Name: ")
url = 'https://timerbackend.geexek.com/rankingBoard.do'
response = requests.post(url, timeout=TIMEOUT, headers={'User-Agent': ua, 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'}, data={
    'rankId': rankId,
    'keyword': keyword,
    'isSingleRank': '1',
    'queryType': '1100',
    'markNo': markNo
})

# TODO: this silently takes the first match - if the name or bib matches several runners (e.g. a common surname,
#       or a blank name with a partial bib) it may track the wrong person. List the matches and ask which one when
#       there is more than one. The search response's status code isn't checked either.
matches = response.json()['list']
if not matches:
    print("Runner not found")
    sys.exit()
runner_id = matches[0]['id']

url = 'https://timerbackend.geexek.com/getScoreInfo.do'
while True:
    try:
        response = requests.post(url, timeout=TIMEOUT, headers={'User-Agent': ua}, data={
            'type': 2,
            'cmptId': cmptId,
            'id': runner_id,
            'rankId': rankId,
            'isWeixin': 1
        })
    except requests.exceptions.RequestException as e:
        print(f"Request failed: {e}")
        sleep(10)
        continue

    if response.status_code != 200:
        print(f"Error: {response.status_code}")
        sleep(10)
        continue

    # Build the whole display before clearing the screen, so an unexpected reply (error page, runner not
    # started yet, missing field...) leaves the last good update on screen and just retries
    try:
        score = response.json()
#        print(json.dumps(score, indent=4, ensure_ascii=False))
        total = score['totalScore']
        header = "\n".join([
            f"Event: {score['cmptName']}",
            f"Race: {score['roadName']}",
            f"Name: {total['name']}",
            f"Bib: {total['markNo']}",
            f"Distance: {total['length'] / 1000}km",
            f"Last CP: {total['state']}",
        ])

        cps = score['memberList'][0]['segs'][0]['cps']
        df = pd.DataFrame(cps)
#        print(json.dumps(cps, indent=4, ensure_ascii=False))

        df['distance'] = df['cpLength'].astype(float) / 1000

        df['cpStopTime_admin_ms70'] = df['cpStopTime_admin_ms70'].replace(0, pd.NA)
        df['passingTime'] = pd.to_datetime(df['cpStopTime_admin_ms70'], unit='ms', errors='coerce', utc=True)
        df['passingTime'] = df['passingTime'].dt.tz_convert('Asia/Shanghai')

        df['predictedTime'] = pd.to_datetime(pd.to_numeric(df['cpStopTimePredict'], errors='coerce'), unit='ms', errors='coerce', utc=True)
        df['predictedTime'] = df['predictedTime'].dt.tz_convert('Asia/Shanghai')

        df['cpIndex'] = pd.to_numeric(df['cpIndex'], errors='coerce') - 1

        table = df[['cpIndex', 'cpName', 'distance', 'passingTime', 'speed', 'cpRank', 'predictedTime']].to_markdown(floatfmt=",.2f", headers=['CP', 'Name', 'Distance', 'Time', 'Speed', 'Rank', 'Predicted'], index=False)
    except (ValueError, KeyError, IndexError, TypeError, AttributeError) as e:
        print(f"{strftime('%H:%M:%S')} Unexpected reply from the timing server ({type(e).__name__}: {e}) - retrying in 10s")
        sleep(10)
        continue

    print("\033[H\033[2J", end="")
    print(header)
    print(table)
    sleep(10)
