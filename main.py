import datetime
import time
import gspread
import requests
import re
from bs4 import BeautifulSoup
from oauth2client.service_account import ServiceAccountCredentials

def get_date_from_user():
    start_date = time.mktime(datetime.datetime(2020, 10, 1).timetuple())  # YYYY-MM-DD
    end_date = time.mktime(datetime.datetime(2021, 10, 22).timetuple())  # YYYY-MM-DD
    return [start_date,end_date]


def get_date_list(start_date, end_date):
    URL = "https://in.investing.com/indices/sensex-historical-data?end_date={}4&st_date={}".format(end_date, start_date)
    r = requests.get(URL, headers={"User-Agent": "Mozilla/5.0"})
    soup = BeautifulSoup(r.content, 'html5lib')
    dates = re.findall(r"<span class=\"text\">([A-z]* [0-9]{2}, [0-9]{4})</span>", str(soup))
    array2Ddates = [dates]
    return array2Ddates


def get_rate_list(URL, start_date, end_date):
    URL = URL + "?end_date={}4&st_date={}".format(end_date, start_date)
    r = requests.get(URL, headers={"User-Agent": "Mozilla/5.0"})
    soup = BeautifulSoup(r.content, 'html5lib')
    soup = soup.findAll("tbody")[1]
    dates = re.findall(r"<span class=\"text\">([A-z]* [0-9]{2}, [0-9]{4})</span>", str(soup))
    rates = re.findall(r"([0-9]*,?[0-9]*\.[0-9]+)</span>", str(soup))
    rates_fin = []
    for i in range(0, len(rates), 4):
        rates_fin.append(rates[i])
    for i in range(len(rates_fin)):
        rates_fin[i] = re.sub(r",", "", rates_fin[i])
        rates_fin[i] = float(rates_fin[i])

    array2Dates = [rates_fin]
    return array2Dates


scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)
sheet = client.open('Stock index Auto update')
sheet_instance = sheet.get_worksheet(0)
start_date = get_date_from_user()[0]
end_date = get_date_from_user()[1]

sheet_instance.clear()
sheet_instance.update('B1', get_date_list(start_date, end_date))
index_dictionary = [['NIFTY-50', 's-p-cnx-nifty'],
                    ['SENSEX', 'sensex'],
                    ['NIFTY Junior', 'cnx-nifty-junior'],
                    ['NIFTY Midcap 100', 'cnx-midcap'],
                    ['NIFTY 500', 's-p-cnx-500'],
                    ['NIFTY Smallcap 100', 'cnx-smallcap'],
                    ['Nifty 200', 'cnx-200'],
                    ['BankNifty', 'bank-nifty'],
                    ['Nifty IT', 'cnx-it'],
                    ['Nifty Finance', 'cnx-finance'],
                    ['Nifty Pharma', 'cnx-pharma'],
                    ['Nifty FMCG', 'cnx-fmcg'],
                    ['Nifty Auto', 'cnx-auto']]
for i in range(len(index_dictionary)):
    URL = "https://in.investing.com/indices/{}-historical-data".format(index_dictionary[i][1])
    sheet_instance.update('A{}'.format(i + 2), index_dictionary[i][0])
    sheet_instance.update('B{}'.format(i + 2), get_rate_list(URL, start_date, end_date))
