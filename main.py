from urllib.request import Request, urlopen
from bs4 import BeautifulSoup
from datetime import date
from win10toast import ToastNotifier
import webbrowser
import lxml

import os.path

from google.auth.transport.requests import Request as requesting
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient import discovery

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1MpdUiDrzyteyCIdCfyEuOGSsdQt2A-DIBtvrzwa7WPY'
SAMPLE_RANGE_NAME = 'A2:C'
# link to the excel file storing data
excel = "https://docs.google.com/spreadsheets/d/1MpdUiDrzyteyCIdCfyEuOGSsdQt2A-DIBtvrzwa7WPY/edit#gid=0"
date_range = 'Sheet1!C2:C'

token_file = 'C:/Users/daemo/PycharmProjects/ScrappingDou/dist/token.json'
credentials_file = 'C:/Users/daemo/PycharmProjects/ScrappingDou/dist/credentials.json'


creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists(token_file):
    creds = Credentials.from_authorized_user_file(token_file, scopes=SCOPES)
# If there are no (valid) credentials available, let the user log in.

if not creds or not creds.valid:
    try:
        os.remove(token_file)
    except:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(requesting())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_file, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_file, 'w') as token:
            token.write(creds.to_json())

service = discovery.build('sheets', 'v4', credentials=creds, static_discovery=False)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range=SAMPLE_RANGE_NAME).execute()


def check_if_repeat():
    result = service.spreadsheets().values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=date_range).execute()
    values = result.get('values', [])
    last_date = values[-1][0]
    return last_date[1::]


def appending(data):
    result_append = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                          range=SAMPLE_RANGE_NAME, valueInputOption='USER_ENTERED',
                                          insertDataOption='INSERT_ROWS', body={"values": data}).execute()
# get current date to compare with posting date
today = date.today()
current_day = today.strftime('%d/%m/%Y')
day_before = int(current_day.split('/')[0])-1
yesterday = f"{day_before}/{str(today).split('-')[1]}/{str(today).split('-')[0]}"


if check_if_repeat() == yesterday:
    toast = ToastNotifier()
    toast.show_toast('Repetition', 'The check has already been performed today',
                     duration=10)
    exit()
# convert months to digits to compare posting date with current date
months = {"січня": "01", "лютого": "02", "березня": "03",
          "квітня": "04", "травня": "05", "червня": "06",
          "липня": "07", "серпня": "08", "вересня": "09",
          "жовтня": "10", "листопада": "11", "грудня": "12"}
# list to store title, site, date
links_dates = []
# link that will get scrapped
my_url = 'https://jobs.dou.ua/first-job/?from=doufp'

# site throws '403 Forbidden'. this allows to get access to the site
req = Request(my_url, headers={'User-Agent': 'Mozilla/5.0'})
web_byte = urlopen(req).read()
webpage = web_byte.decode('utf-8')

try:
    soup = BeautifulSoup(webpage, 'lxml')
except:
    soup = BeautifulSoup(webpage, "html.parser")


# to convert month into digits
def month_to_num(month):
    for key, value in months.items():
        if key == month[1]:
            posting_date = f'{month[0]}/{value}/{month[2]}'
            return posting_date


# to get titles, links and dates
def get_links_titles():
    count_current_posts = 0
    # to get hot titles, links, dates from a webpage
    for element in soup.find_all('li', class_='l-vacancy'):
        hot = element.find('a', class_='vt')['href']
        title = element.find('a', class_='vt')
        # hot link doesn't have date, returns None-type
        # get its date from the link
        # for hot links only(they are structured differently)
        if 'list_hot' in hot.split('?')[-1]:
            req_hot = Request(hot, headers={'User-Agent': 'Mozilla/5.0'})
            web_byte_hot = urlopen(req_hot).read()
            webpage_hot = web_byte_hot.decode('utf-8')
            try:
                soup_hot = BeautifulSoup(webpage_hot, 'lxml')
            except:
                soup_hot = BeautifulSoup(webpage_hot, "html.parser")
            link_to_hot = soup_hot.find('div', class_='date')
            month_hot = link_to_hot.text[:link_to_hot.text.index(yesterday.split('/')[2]) + 4].strip().split(' ')
            posting_date_hot = month_to_num(month_hot)
            # dont add data if it's an old vacancy
            if posting_date_hot != yesterday:
                pass
            # add data if it's a new vacancy(from today)
            elif posting_date_hot == yesterday:
                count_current_posts += 1
                links_dates.append([title.text, hot, posting_date_hot])
        else:
            post_date = element.find('div', class_='date')
            month = post_date.text.split(' ')
            posting_date = month_to_num(month)
            if posting_date != yesterday:
                pass
            elif posting_date == yesterday:
                count_current_posts += 1
                links_dates.append([title.text, hot, posting_date])
    return count_current_posts


# open notifications. no vacancies or there are vacancies
def notification_open(count_current_posts):
    toast = ToastNotifier()
    try:
        if count_current_posts == 0:  # to tell if there are no posts from today
            toast.show_toast('No new vacancies', f'{yesterday} we got no new vacancies yesterday', duration=10,
                             icon_path='C:\\Users\\daemo\\PycharmProjects\\ScrappingDou\\icon-console.ico')
        # throw a notification an open excel if there were vacancies from today
        elif count_current_posts > 0:
            open_excel = webbrowser.open_new(excel)
            toast.show_toast('Number of vacancies', f'Yesterday - {yesterday} we got {count_current_posts} new vacancies',
                             duration=10, icon_path='C:\\Users\daemo\\PycharmProjects\\ScrappingDou\\icon-console.ico')
    except:
        if count_current_posts == 0:  # to tell if there are no posts from today
            toast.show_toast('No new vacancies', f'{yesterday} we got no new vacancies yesterday', duration=10,
                             icon_path='')
        # throw a notification an open excel if there were vacancies from today
        elif count_current_posts > 0:
            open_excel = webbrowser.open_new(excel)
            toast.show_toast('Number of vacancies', f'Yesterday - {yesterday} we got {count_current_posts} new vacancies',
                             duration=10, icon_path='')


print('here')
get_links_titles()
appending(links_dates)
notification_open(get_links_titles())
