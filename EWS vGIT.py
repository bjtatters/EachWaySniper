# region - IMPORT MODULES
import PySimpleGUI as sg 
import datetime
from numpy.core.numeric import NaN
import requests
from itertools import groupby
import pandas as pd
import numpy as np
import time
from requests import get
from bs4 import BeautifulSoup
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import schedule
import sys
import openpyxl
from openpyxl import Workbook
import os
import traceback
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
pd.options.mode.chained_assignment = None
# endregion
# region - GUI
sg.ChangeLookAndFeel('Purple')
layout = [
    [sg.Text('Each Way Sniper', font=('Helvetica', 20), justification='left')],
    [sg.Text('_'  * 100, size=(65, 1))],
    [sg.Text('Choose Desired EWS Stake, Extra Place Criteria (standard = 8) and Desired Extra Place Stake', font=('Helvetica', 15), justification='left')],
    [sg.Text('Stake', size=(15, 1)), sg.InputText('2')],[sg.Text('Criteria', size=(15, 1)), sg.InputText('8')],[sg.Text('Stake', size=(15, 1)), sg.InputText('2')],
    [sg.Text('_'  * 100, size=(65, 1))],
    [sg.Text('Choose Minimum Desired ROI (%) and Betfair Commission (standard = 0.05)', font=('Helvetica', 15), justification='left')],
    [sg.Text('Minimum ROI', size=(15, 1)), sg.InputText('0')],[sg.Text('Commission', size=(15, 1)), sg.InputText('0.05')],
    [sg.Text('_'  * 100, size=(65, 1))],
    [sg.Text('Input Email', font=('Helvetica', 15), justification='left')],
    [sg.Text('Email', size=(15, 1)), sg.InputText('')],
    [sg.Text('_'  * 100, size=(65, 1))],
    [sg.Text('Choose Output Folders', font=('Helvetica', 15), justification='left')], 
    [sg.Text('EWS Folder', size=(15, 1), auto_size_text=False),sg.InputText(''), sg.FolderBrowse()],\
    [sg.Text('Extra Place Folder', size=(15, 1), auto_size_text=False),sg.InputText(''), sg.FolderBrowse()],\
    [sg.Text('Odds Today Folder', size=(15, 1), auto_size_text=False),sg.InputText(''), sg.FolderBrowse()],\
    [sg.Text('Horse Check Folder', size=(15, 1), auto_size_text=False),sg.InputText(''), sg.FolderBrowse()],\
    [sg.Text('Errors Folder', size=(15, 1), auto_size_text=False),sg.InputText(''), sg.FolderBrowse()],
    [sg.Text('_'  * 100, size=(65, 1))],
    [sg.Text('Timeframe (in hours) and Program Repetition Frequency (in seconds)', font=('Helvetica', 15), justification='left')],
    [sg.Text('Time', size=(15, 1)), sg.InputText('1')], [sg.Text('Repetition Frequency', size=(15, 1)), sg.InputText('600')],
    [sg.Text('_'  * 100, size=(65, 1))],
    [sg.OK()]
]
window = sg.Window('EWS', layout, font=("Helvetica", 12))      
event, values = window.read()
window.close()
# endregion
# region - RACES TODAY 
SCRAPINGBEE_API_KEY = "" # to avoid getting blocked by Betfair / Oddschecker, I use ScrapingBee. You can use their own 14-day free trial and enter your API key here.
endpoint = "https://app.scrapingbee.com/api/v1"

def racestoday():
    
    params = {
            'api_key': SCRAPINGBEE_API_KEY,
            'url': 'https://apieds.betfair.com/api/eds/meeting-races/v4?_ak=nzIFcwyWhrlwYMrh&countriesGroup=%5B%5B%22GB%22,%22IE%22%5D%5D&eventTypeId=7&marketStartingAfter={}-{}-{}T00:00:00.000Z&marketStartingBefore={}-{}-{}T23:59:59.999Z'.format(today.year,month,day,today.year,month,day)
    }

    content = str(requests.get(endpoint, params=params).content)
    
    marketId_list, raceId_list, starttime_list, venue_list = ([] for i in range(4)) 
    for i, _ in enumerate(content):
        needle1 = "marketId"
        needle2 = "raceId"
        needle3 = "startTime"
        needle4 = '"venue"'
        if content[i:i + len(needle1)] == needle1:
            marketId_list.append(content[i+11:i+22])
        if content[i:i + len(needle2)] == needle2:
            raceId_list.append(content[i+9:i+22])
        if content[i:i + len(needle3)] == needle3:
            starttime_list.append(content[i+23:i+28])
        if content[i:i + len(needle4)] == needle4:
            starttime_list.append(content[i+9:i+25].split('"',1)[0])
        if content[i:i + len(needle4)] == needle4:
            venue_list.append(content[i+9:i+25].split('"',1)[0])
        
    venue_indices = []
    for venue in venue_list:
        venue_index = starttime_list.index(venue)
        venue_indices.append(venue_index)

    race_dict = {}
    if len(venue_list) == 1 and len(starttime_list) > 0: # be aware of this change to 0 from 2
        dic1 = dict.fromkeys(starttime_list[0:(int(venue_indices[0]))], venue_list[0])
        race_dict.update(dic1)

    if len(venue_indices) > 1:
        dic2 = dict.fromkeys(starttime_list[0:(int(venue_indices[0]))], venue_list[0])
        race_dict.update(dic2)
        for i in range(1,len(venue_indices)):
            dic3 = dict.fromkeys(starttime_list[int(venue_indices[i-1])+1:(int(venue_indices[i]))], venue_list[i])
            race_dict.update(dic3)

    global races_today, races_today_rows 
    races_today = pd.DataFrame()
    races_today['Location'] = race_dict.values()
    races_today['Time'] = race_dict.keys()
    races_today['MarketId'] = marketId_list
    races_today['RaceId'] = raceId_list

    global v_list, t_list, m_list, r_list, errors
    time_period = int(values[11])*100
    t_list, v_list, m_list, r_list, errors = ([] for i in range(5)) 
    races_today['Time'] =  pd.to_datetime(races_today['Time']) # format='%H:%M')
    for i in range(0,len(races_today)):
        delta = races_today.iloc[i]['Time'] - datetime.datetime.now()
        if delta.seconds < (int(values[11])*3600):
            t_list.append(races_today.iloc[i]['Time'])
            v_list.append(races_today.iloc[i]['Location'])
            m_list.append(races_today.iloc[i]['MarketId'])
            r_list.append(races_today.iloc[i]['RaceId'])

    races_today = pd.DataFrame({'Location':v_list,'Time':t_list,'MarketId':m_list,'RaceId':r_list})
    print(races_today)
    races_today_rows = races_today.values.tolist()
# endregion
# region - BETFAIR AND ODDS TODAY - need to find a way to account for missing odds in the first 6 horses
def betfair():
    global odds_today
    odds_today = pd.DataFrame(columns = ['Location','Time', 'MarketId', 'RaceId', 'Bet1', 'Lay1', 'Bet2', 'Lay2', 
    'Bet3', 'Lay3', 'Bet4', 'Lay4', 'Bet5', 'Lay5', 'Bet6', 'Lay6','Place Bet1', 'Place Lay1', 'Place Bet2', 'Place Lay2', 'Place Bet3',
    'Place Lay3', 'Place Bet4', 'Place Lay4', 'Place Bet5', 'Place Lay5', 'Place Bet6', 'Place Lay6', 'Horse1', 'Horse2', 'Horse3', 
    'Horse4', 'Horse5', 'Horse6', 'WIN Data', 'PLACE Data', 'PLACE Link', 'WIN Link', 'Places'])
    
    def betfair_odds(MarketId, RaceId): 
        RaceId2 = RaceId[0:8]
        params = {
            'api_key': SCRAPINGBEE_API_KEY,
            'url': 'https://apieds.betfair.com/api/eds/racing-navigation/v1?_ak=nzIFcwyWhrlwYMrh&eventTypeId=7&meetingId={}&navigationType=racesbymeeting&raceId={}'.format(RaceId2,RaceId)
        }

        content_fp = str(requests.get(endpoint, params=params).content)
        
        marketId_list2, market_type_set = ([] for i in range(2))
        needle5 = "marketId"
        needle6 = '"marketType":"'
        for i, _ in enumerate(content_fp):
            if content_fp[i:i + len(needle5)] == needle5:
                marketId_list2.append(content_fp[i+11:i+22])
            if content_fp[i:i + len(needle6)] == needle6:
                market_type_set.append(content_fp[i+14:i+30].split('"',1)[0])

        if 'PLACE' not in market_type_set:
            errors.append(MarketId)
            print('ERROR: NO PLACE MARKET FOUND FOR {}'.format(MarketId))
            return
        if 'PLACE' in market_type_set:
            place_id = marketId_list2[market_type_set.index('PLACE')] 
            
            params = {
                'api_key': SCRAPINGBEE_API_KEY,
                'url': 'https://ero.betfair.com/www/sports/exchange/readonly/v1/bymarket?_ak=nzIFcwyWhrlwYMrh&alt=json&currencyCode=GBP&locale=en_GB&marketIds={}&rollupLimit=10&rollupModel=STAKE&types=MARKET_STATE,MARKET_RATES,MARKET_DESCRIPTION,EVENT,RUNNER_DESCRIPTION,RUNNER_STATE,RUNNER_EXCHANGE_PRICES_BEST,RUNNER_METADATA,MARKET_LICENCE,MARKET_LINE_RANGE_INFO'.format(str(MarketId))
            }

            win_content = str(requests.get(endpoint, params=params).content)

            params = {
                'api_key': SCRAPINGBEE_API_KEY,
                'url': 'https://ero.betfair.com/www/sports/exchange/readonly/v1/bymarket?_ak=nzIFcwyWhrlwYMrh&alt=json&currencyCode=GBP&locale=en_GB&marketIds={}&rollupLimit=10&rollupModel=STAKE&types=MARKET_STATE,MARKET_RATES,MARKET_DESCRIPTION,EVENT,RUNNER_DESCRIPTION,RUNNER_STATE,RUNNER_EXCHANGE_PRICES_BEST,RUNNER_METADATA,MARKET_LICENCE,MARKET_LINE_RANGE_INFO'.format(str(place_id))
            }

            place_content = str(requests.get(endpoint, params=params).content)

            win_url = 'https://ero.betfair.com/www/sports/exchange/readonly/v1/bymarket?_ak=nzIFcwyWhrlwYMrh&alt=json&currencyCode=GBP&locale=en_GB&marketIds={}&rollupLimit=10&rollupModel=STAKE&types=MARKET_STATE,MARKET_RATES,MARKET_DESCRIPTION,EVENT,RUNNER_DESCRIPTION,RUNNER_STATE,RUNNER_EXCHANGE_PRICES_BEST,RUNNER_METADATA,MARKET_LICENCE,MARKET_LINE_RANGE_INFO'.format(str(MarketId))
            place_url = 'https://ero.betfair.com/www/sports/exchange/readonly/v1/bymarket?_ak=nzIFcwyWhrlwYMrh&alt=json&currencyCode=GBP&locale=en_GB&marketIds={}&rollupLimit=10&rollupModel=STAKE&types=MARKET_STATE,MARKET_RATES,MARKET_DESCRIPTION,EVENT,RUNNER_DESCRIPTION,RUNNER_STATE,RUNNER_EXCHANGE_PRICES_BEST,RUNNER_METADATA,MARKET_LICENCE,MARKET_LINE_RANGE_INFO'.format(str(place_id))
            
            win_odds_list, names_list, place_odds_list, place_winners = ([] for i in range(4))
            needle7 = '"price"'
            needle8 = '"runnerName":"'
            needle9 = '"price"'
            needle10 = '"numberOfWinners"'
            for i, _ in enumerate(win_content):
                if win_content[i:i + len(needle7)] == needle7:
                    win_odds_list.append(win_content[i+8:i+12])
                if win_content[i:i + len(needle8)] == needle8:
                    names_list.append(win_content[i+14:i+44].split('"',1)[0])
            for x, _ in enumerate(place_content):
                if place_content[x:x + len(needle9)] == needle9:
                    place_odds_list.append(place_content[x+8:x+12])
            for w, _ in enumerate(place_content):
                if place_content[w:w + len(needle10)] == needle10:
                    place_winners.append(place_content[w+18:w+19])

            if len(win_odds_list) < 24 or len(place_odds_list) < 24 or len(names_list)<4:
                errors.append(MarketId)
                print('ERROR: LESS THAN 4 HORSES FOUND FOR {}'.format(MarketId))
                return
            if len(win_odds_list) >= 24 and len(place_odds_list) >= 24 and len(names_list) >=4:
                row_base = races_today_rows[m_list.index(MarketId)]
                if len(names_list) >= 6 and len(win_odds_list) >= 34 and len(place_odds_list) >= 34: # changed from 36 to 34
                    names_num = 36 
                else:
                    names_num = int(6 * round((min([len(win_odds_list), len(place_odds_list)]))/6))      
                for i in range(0,names_num,3):
                    row_base.append(str(win_odds_list[i]))
                if len(row_base) != 16:
                    n = 16 - len(row_base)
                    row_base += n * ['NO ODDS']
                for i in range(0,names_num,3):
                    row_base.append(str(place_odds_list[i]))
                if len(row_base) != 28:
                    n = 28 - len(row_base)
                    row_base += n * ['NO ODDS']
                for i in range(0,int(names_num/6)):
                    row_base.append(str(names_list[i]))
                if len(row_base) != 34:
                    n = 34 - len(row_base)
                    row_base += n * ['NO NAME']
                row_base.append(place_url)
                row_base.append(win_url)
                row_base.append('https://www.betfair.com/exchange/plus/horse-racing/market/{}'.format(MarketId))
                row_base.append('https://www.betfair.com/exchange/plus/horse-racing/market/{}'.format(place_id))
                row_base.append(place_winners[0])
                row_base_refined = []
                for val in row_base:
                    if row_base.index(val) < 4 or row_base.index(val) > 27:
                        row_base_refined.append(val)
                    else:
                        if val[-1] != "," and val != 'NO ODDS':
                            row_base_refined.append(val)
                        if val[-1] == ",":
                            row_base_refined.append(val[:-1])
                        if val == 'NO ODDS':
                            row_base_refined.append(val)
                if len(row_base_refined) == len(row_base):
                    odds_today.loc[len(odds_today)] = row_base_refined
                else:
                    print('ERROR: ROW BASE AND ROW BASE REFINED ARE DIFFERENT LENGTHS FOR {}'.format(MarketId))
                    print(row_base)
                    print(row_base_refined)
                    return

    for i in range(0,len(m_list)):
        betfair_odds(m_list[i],r_list[i])
    
    print('LENGTH OF ODDS TODAY DATASET: {} @ {}'.format(len(odds_today),runtime))
    odds_today.to_excel('{}/Odds-Today-{}-{}-{}@{}.xlsx'.format(values[8],today.year,month,day,runtime))
# endregion
# region - ODDSCHECKER 1
def oddschecker_1():
    global not_4_horse_race
    horse_class, horse_bk, horse_ew_denom, horse_ew_places, horse_fodds, horse_o, horse_odig, horse_names, not_4_horse_race, check_errors, horse_check_list = ([] for i in range(11))
    market_id_list = odds_today['MarketId'].tolist()
    extension_lists = odds_today.values.tolist()

    def oddschecker_1_sub(venue,time,MarketId,stake):
        odds_row = market_id_list.index(MarketId)
        venue = venue.replace(' ','-')
        race_url = "https://www.oddschecker.com/horse-racing/{}-{}-{}-{}/{}:{}/winner".format(today.year,month,day,venue.lower(),time.hour,time.minute)

        params = {
            'api_key': SCRAPINGBEE_API_KEY,
            'url': "https://www.oddschecker.com/horse-racing/{}-{}-{}-{}/{}:{}/winner".format(today.year,month,day,venue.lower(),time.hour,time.minute)
        }

        race_response = requests.get(endpoint, params=params)
        race_soup = BeautifulSoup(race_response.content, 'html.parser')

        horse_odds = race_soup.find_all('td',class_='bc')
        horse_names = race_soup.find_all('a', {'class':['popup selTxt has-tip', 'popup selTxt']})
        names, positions = ([] for i in range(2))
        for x in range(0, len(horse_names)):
            name = str(horse_names[x])
            positions.append(x+1)
            needle18 = 'data-name="'
            for i, _ in enumerate(name):
                try:
                    if name[i:i + len(needle18)] == needle18:
                        names.append(name[i+11:i+41].split('"',1)[0])
                except:
                    names.append(np.nan)
        name_dict = dict(zip(positions,names))
        
        horse_bk_index = []   
        for x in range(0, len(horse_odds)): 
            horse = horse_odds[x]
            horse_a = str(horse)
            needle12 = 'data-bk="'
            for i, _ in enumerate(horse_a):
                try:
                    if horse_a[i:i + len(needle12)] == needle12:
                        horse_bk_index.append(horse_a[i+9:i+13].split('"',1)[0])
                except:
                    horse_bk_index.append(np.nan)
        indices = [i for i, x in enumerate(horse_bk_index) if x == "B3"]
        if len(indices)<4:
            not_4_horse_race.append(MarketId)
        else:
            indices.append(indices[(len(indices)-1)] + indices[1])

        for x in range(0, len(horse_odds)):
            horse_check = []
            for i in range(0, len(indices)-1):
                if x >= int(indices[i]) and x < int(indices[i+1]):
                    horse_check.append(i+1)
            if len(horse_check) == 0:
                horse_check.append(np.nan)
            horse = str(horse_odds[x])
            needle11 = 'class="'
            needle12 = 'data-bk="'
            needle13 = 'data-ew-denom="'
            needle14 = 'data-ew-places="'
            needle15 = 'data-fodds="'
            needle16 = 'data-o="'
            needle17 = 'data-odig="'
            needle_list = [needle11,needle12,needle13,needle14,needle15,needle16,needle17]
            character_list = [7,9,15,16,12,8,11]
            append_list = [horse_class, horse_bk, horse_ew_denom, horse_ew_places, horse_fodds, horse_o, horse_odig]
            for i, _ in enumerate(horse):
                for n in range(0, len(needle_list)):
                    try:
                        if horse[i:i + len(needle_list[n])] == needle_list[n]:
                            append_list[n].append(horse[i+character_list[n]:i+25].split('"',1)[0])
                            horse_check.append(horse[i+character_list[n]:i+25].split('"',1)[0])
                    except:
                        append_list[n].append(np.nan)
                        horse_check.append(np.nan)
            try:
                horse_check.append(name_dict.get(horse_check[0])) 
            except:
                horse_check.append(np.nan)
            horse_check.append(race_url)
            if len(horse_check) == 10: 
                horse_check.extend(extension_lists[odds_row])
                horse_check_list.append(horse_check)
            if len(horse_check) < 10 or np.nan in horse_check == True:
                check_errors.append('{} @ {}'. format(venue, time))

    for i in range(0,len(odds_today)):
        oddschecker_1_sub(odds_today.iloc[i]['Location'],odds_today.iloc[i]['Time'],odds_today.iloc[i]['MarketId'],values[0])

    global horse_check_df, check_df_updated
    horse_check_df = pd.DataFrame.from_records(horse_check_list, columns = ['Oddschecker Number','Class','BK','EW-Denom','EW-Places',
    'Fodds','O','Odig','Name','Oddschecker Link','Location','Time', 'MarketId', 'RaceId', 'Bet1', 'Lay1', 'Bet2', 'Lay2', 
    'Bet3', 'Lay3', 'Bet4', 'Lay4', 'Bet5', 'Lay5', 'Bet6', 'Lay6','Place Bet1', 'Place Lay1', 'Place Bet2', 'Place Lay2', 'Place Bet3',
    'Place Lay3', 'Place Bet4', 'Place Lay4', 'Place Bet5', 'Place Lay5', 'Place Bet6', 'Place Lay6', 'Horse1', 'Horse2', 'Horse3', 
    'Horse4', 'Horse5', 'Horse6', 'WIN Data', 'PLACE Data', 'PLACE Link', 'WIN Link', 'Places'])
    print('LENGTH OF HORSE CHECK DF: {}'.format(len(horse_check_df)))
    horse_check_df.to_excel('{}/Horse-Check-Raw-{}-{}-{}@{}.xlsx'.format(values[9],today.year,month,day,runtime))
    numbers = [1,2,3,4,5,6]
    check_df_updated = horse_check_df[horse_check_df['Oddschecker Number'].isin(numbers)]
    check_df_updated = check_df_updated[~check_df_updated['MarketId'].isin(not_4_horse_race)]
    check_df_updated = check_df_updated[check_df_updated['O'] != 'SP']
    check_df_updated = check_df_updated.drop(columns = ['Class','Fodds','MarketId','WIN Data', 'PLACE Data','Odig'])
     
    print('LENGTH OF HORSE CHECK UPDATED DF: {}'.format(len(check_df_updated)))
 
# endregion 
# region - ODDSCHECKER 2
def oddschecker_2():
    stake_list = [int(values[0])] * len(check_df_updated)
    check_df_updated['Stake'] = stake_list
    horse_check_o = check_df_updated['O'].tolist()

    def convert_to_float(frac_str):
        try:
            return float(frac_str)
        except ValueError:
            num, denom = frac_str.split('/')
            try:
                leading, num = num.split(' ')
                whole = float(leading)
            except ValueError:
                whole = 0
            frac = float(num) / float(denom)
            return (whole - frac) if whole < 0 else (whole + frac)
    intermediate = []    
    for o in horse_check_o: # there was one error here - be careful of any exceptions that may arise
        if len(o) > 2:
            if o[1] == '/':
                w = convert_to_float(o)
                intermediate.append(w)
            if o[2] == '/':
                x = convert_to_float(o)
                intermediate.append(x)   
            if len(o)>=4 and o[3] == '/':
                y = convert_to_float(o)
                intermediate.append(y)
        if len(o) <= 2:
            intermediate.append(o)
    o_converted = []
    for flo in intermediate:
        if isinstance(flo,float) == True:
            o_converted.append(flo)
        if isinstance(flo,float) == False:
            o_converted.append(int(flo))
    
    check_df_updated['O'] = o_converted
# endregion
# region - FMC 1
def fmc_1():
    betfair_numbers, num_errors, extra_horse_drop = ([] for i in range(3))
    for i in range(0,len(check_df_updated)):
        if check_df_updated.iloc[i]['Name'] == check_df_updated.iloc[i]['Horse1']:
            betfair_numbers.append('1')
        if check_df_updated.iloc[i]['Name'] == check_df_updated.iloc[i]['Horse2']:
            betfair_numbers.append('2')
        if check_df_updated.iloc[i]['Name'] == check_df_updated.iloc[i]['Horse3']:
            betfair_numbers.append('3')
        if check_df_updated.iloc[i]['Name'] == check_df_updated.iloc[i]['Horse4']:
            betfair_numbers.append('4')
        if check_df_updated.iloc[i]['Name'] == check_df_updated.iloc[i]['Horse5']:
            betfair_numbers.append('5')
        if check_df_updated.iloc[i]['Name'] == check_df_updated.iloc[i]['Horse6']:
            betfair_numbers.append('6')
        if check_df_updated.iloc[i]['Name'] != check_df_updated.iloc[i]['Horse1'] and check_df_updated.iloc[i]['Name'] != check_df_updated.iloc[i]['Horse2'] and check_df_updated.iloc[i]['Name'] != check_df_updated.iloc[i]['Horse3'] and check_df_updated.iloc[i]['Name'] != check_df_updated.iloc[i]['Horse4'] and check_df_updated.iloc[i]['Name'] != check_df_updated.iloc[i]['Horse5'] and check_df_updated.iloc[i]['Name'] != check_df_updated.iloc[i]['Horse6']:
            betfair_numbers.append(np.nan)
            num_errors.append(i)

    print('CHECK: {} HORSES WHICH ARE IN TOP 6 OF ODDSCHECKER BUT NOT TOP 6 OF BETFAIR EXCHANGE'.format(len(num_errors)))
    check_df_updated['Betfair Horse Number'] = betfair_numbers
    global fmc, commission 
    fmc = check_df_updated.dropna()
    print('CHECK: {} NaN VALUES'.format(len(check_df_updated) - len(fmc)))
    is_NaN = check_df_updated.isnull()
    errors = check_df_updated[is_NaN.any(axis=1)]
    errors.to_excel('{}/Errors-{}-{}-{}@{}.xlsx'.format(values[10],today.year,month,day,runtime))
    fmc.to_excel('{}/Horse-Check-Cleaned-{}-{}-{}@{}.xlsx'.format(values[9],today.year,month,day,runtime))

    for i in range(0, len(fmc)):
        x = fmc.iloc[i]['Betfair Horse Number']
        if fmc.iloc[i]['Bet{}'.format(x)] == 'NO ODDS' or fmc.iloc[i]['Lay{}'.format(x)] == 'NO ODDS' or fmc.iloc[i]['Place Bet{}'.format(x)] == 'NO ODDS' or fmc.iloc[i]['Place Lay{}'.format(x)] == 'NO ODDS':
            extra_horse_drop.append(i)
    fmc = fmc.drop(extra_horse_drop) 

    float_conversion_list = ['Bet1', 'Lay1', 'Bet2', 'Lay2', 'Bet3', 'Lay3', 'Bet4', 'Lay4', 'Bet5', 'Lay5', 'Bet6', 'Lay6',
    'Place Bet1', 'Place Lay1', 'Place Bet2', 'Place Lay2', 'Place Bet3', 'Place Lay3', 'Place Bet4', 'Place Lay4', 'Place Bet5', 
    'Place Lay5', 'Place Bet6', 'Place Lay6']
    for lis in float_conversion_list:
        fmc[lis] = fmc[lis].astype(float)
    int_conversion_list = ['Stake', 'EW-Denom'] #, 'Betfair Horse Number']
    for lis in int_conversion_list:
        fmc[lis] = fmc[lis].astype(int)

    commission = float(values[4])
    placeodds4 = 0.75
    placeodds5 = 0.8

    place_odds_list = []
    for i in range(0,len(fmc.loc[:,'Location'])): 
        if fmc.iloc[i]['EW-Denom'] == 4 or fmc.iloc[i]['EW-Denom'] == 5:
            if fmc.iloc[i]['EW-Denom'] == 5:
                place_odds_5 = placeodds5 + ((placeodds5/4)*float(fmc.iloc[i]['O']))
                place_odds_list.append(float(place_odds_5))
            if fmc.iloc[i]['EW-Denom'] == 4:
                place_odds_4 = placeodds4 + ((placeodds4/3)*float(fmc.iloc[i]['O']))
                place_odds_list.append(float(place_odds_4))
        else: # if fmc.iloc[i]['EW-Denom'] == 1 or fmc.iloc[i]['EW-Denom'] == 0 or fmc.iloc[i]['EW-Denom'] == 'N/A' or fmc.iloc[i]['EW-Denom'] == 2 or fmc.iloc[i]['EW-Denom'] == 3 or fmc.iloc[i]['EW-Denom'] == 6:
            place_odds_list.append(np.nan)
    fmc['Place Odds'] = place_odds_list

    wb_laystake_list, wb_liability_list, wb_bb_win, wb_lb_win, pb_laystake_list, \
    pb_liability_list, pb_bb_win, pb_lb_win = ([] for i in range(8))
    calculations_list = [wb_laystake_list, wb_liability_list, wb_bb_win, wb_lb_win, pb_laystake_list, pb_liability_list, pb_bb_win, pb_lb_win]
    for i in range(0,len(fmc)):    
        num = int(fmc.iloc[i]['Betfair Horse Number'])
        wb_laystake = (fmc.iloc[i]['O']*fmc.iloc[i]['Stake'])/(fmc.iloc[i]['Lay{}'.format(str(num))]-commission)
        wb_laystake_list.append(wb_laystake)
        wb_liability = (fmc.iloc[i]['Lay{}'.format(str(num))]-1)*wb_laystake
        wb_liability_list.append(wb_liability)
        wb_bb_wins = (fmc.iloc[i]['Stake']*(fmc.iloc[i]['O']-1))-wb_liability
        wb_bb_win.append(wb_bb_wins)
        wb_lb_wins = (wb_laystake*(1-commission))-fmc.iloc[i]['Stake']
        wb_lb_win.append(wb_lb_wins)
        pb_laystake = (fmc.iloc[i]['Place Odds']*fmc.iloc[i]['Stake'])/(fmc.iloc[i]['Place Lay{}'.format(str(num))]-commission)
        pb_laystake_list.append(pb_laystake)
        pb_liability = (fmc.iloc[i]['Place Lay{}'.format(str(num))]-1)*pb_laystake
        pb_liability_list.append(pb_liability)
        pb_bb_wins = (fmc.iloc[i]['Stake']*(fmc.iloc[i]['Place Odds']-1))-pb_liability
        pb_bb_win.append(pb_bb_wins)
        pb_lb_wins = (pb_laystake*(1-commission))-fmc.iloc[i]['Stake']
        pb_lb_win.append(pb_lb_wins)

    calculations_dict_1 = {'WB Laystake':wb_laystake_list,'WB Liability':wb_liability_list, 'WB BB Win':wb_bb_win, 'WB LB Win':wb_lb_win,
    'PB Laystake':pb_laystake_list,'PB Liability':pb_liability_list, 'PB BB Win':pb_bb_win, 'PB LB Win':pb_lb_win}
    for key, value in calculations_dict_1.items():
        fmc[key] = value
    for key in calculations_dict_1:
        fmc[key] = fmc[key].astype(float)

    total_stake, total_liability, horse_wins, horse_loses, roi_horse_wins, roi_horse_loses, place_drop = ([] for i in range(7))
    for i in range(0,len(fmc)):
        tstake = (fmc.iloc[i]['Stake'])*2
        total_stake.append(tstake)
        tliability = fmc.iloc[i]['WB Liability']+fmc.iloc[i]['PB Liability']
        total_liability.append(tliability)
        hw = fmc.iloc[i]['WB BB Win'] + fmc.iloc[i]['PB BB Win']
        horse_wins.append(hw)
        hl = fmc.iloc[i]['WB LB Win'] + fmc.iloc[i]['PB LB Win']
        horse_loses.append(hl)
        roihw = (hw/tstake)*100
        roi_horse_wins.append(roihw)
        roihl = (hl/tstake)*100
        roi_horse_loses.append(roihl)

    calculations_dict_2 = {'Total Stake':total_stake,'Total Liability':total_liability,'Profit (Horse Wins)':horse_wins,'Profit (Horse Loses)':horse_loses,
    'ROI_(Horse_Wins)':roi_horse_wins,'ROI_(Horse_Loses)':roi_horse_loses}
    for key, value in calculations_dict_2.items():
        fmc[key] = value

    fmc = fmc.dropna()
    fmc = fmc.sort_values(['ROI_(Horse_Wins)'], ascending = False)

# endregion
# region - FMC FIFTH
def fmcfifth():
    global fmc_fifth
    fmc_fifth = fmc[(fmc['Places'] < fmc['EW-Places'])]

    for i in range(0,len(fmc_fifth)):
        fmc_fifth.iloc[i]['Stake'] == float(values[2])
    global sw_list, sp_in_sp_list, sp_in_ap_list, sp_dp_list, met_list
    sw_list, sp_in_sp_list, sp_in_ap_list, sp_dp_list, met_list = ([] for i in range(5))
    for i in range(0,len(fmc_fifth)):
        num = fmc_fifth.iloc[i]['Betfair Horse Number']
        sw = ((fmc_fifth.iloc[i]['O']-1)*fmc_fifth.iloc[i]['Stake']) - ((fmc_fifth.iloc[i]['Lay{}'.format(num)]-1)*fmc_fifth.iloc[i]['WB Laystake']) + ((fmc_fifth.iloc[i]['Place Odds']-1)*fmc_fifth.iloc[i]['Stake']) - ((fmc_fifth.iloc[i]['Place Lay{}'.format(num)]-1)*fmc_fifth.iloc[i]['PB Laystake'])
        sw_list.append(sw)
        sp_in_sp = -fmc_fifth.iloc[i]['Stake'] + (fmc_fifth.iloc[i]['WB Laystake']*(1-commission)) + ((fmc_fifth.iloc[i]['Place Odds']-1)*fmc_fifth.iloc[i]['Stake']) - ((fmc_fifth.iloc[i]['Place Lay{}'.format(num)]-1)*fmc_fifth.iloc[i]['PB Laystake'])
        sp_in_sp_list.append(sp_in_sp)
        sp_in_ap = -fmc_fifth.iloc[i]['Stake'] + (fmc_fifth.iloc[i]['WB Laystake']*(1-commission)) + ((fmc_fifth.iloc[i]['Place Odds']-1)*fmc_fifth.iloc[i]['Stake']) + (fmc_fifth.iloc[i]['PB Laystake']*(1-commission))
        sp_in_ap_list.append(sp_in_ap)
        sp_dp = -fmc_fifth.iloc[i]['Stake'] + (fmc_fifth.iloc[i]['WB Laystake']*(1-commission)) - fmc_fifth.iloc[i]['Stake'] + (fmc_fifth.iloc[i]['PB Laystake']*(1-commission))
        sp_dp_list.append(sp_dp)
        met = -(sp_in_ap)/((sw + sp_in_sp + sp_dp)/3)
        met_list.append(met)

    if len(sw_list) !=  len(fmc_fifth):
        print('LENGTH ERROR FOR EXTRA PLACE DATASET')

    calculations_dict_3 = {'SW':sw_list, 'SP in SP':sp_in_sp_list,'SP in AP':sp_in_ap_list,'SP DP':sp_dp_list,'Judge':met_list}
    for key, value in calculations_dict_3.items():
        fmc_fifth[key] = value
# endregion
# region - FMC 2
def fmc_2():
    print('LENGTH OF EWS DATASET: {} @ {}'.format(len(fmc),runtime))
    target_ROI = int(values[3])
    target_judge = int(values[1])
    global fmc_fifth

    fmc_fifth.sort_values(['Judge'], ascending = False)
    fmc_fifth.to_excel('{}/Extra-Place-{}-{}-{}@{}.xlsx'.format(values[7],today.year,month,day,runtime))
    print('LENGTH OF EXTRA PLACE DATASET: {} @ {}'.format(len(fmc_fifth),runtime))
    fmc_fifth = fmc_fifth.query('`Judge` >= @target_judge')
    print('NUM OF PROFITABLE EXTRA PLACE BETS (Judge > {}): {} @ {}'.format(str(target_judge),len(fmc_fifth),runtime))
    fmc.to_excel('{}/EWS-{}-{}-{}@{}.xlsx'.format(values[6],today.year,month,day,runtime))

    global fmc_current, fmc_bookie, fmc_current_bookie
    fmc_current = fmc.query('`ROI_(Horse_Wins)` >= @target_ROI')
    print('NUM OF PROFITABLE BETS (ANY BOOKIE): {} @ {}'.format(len(fmc_current),runtime))
    fmc_bookie = fmc_current[fmc_current['BK'].isin(['B3','SK','PP'])] # expand on the bookies included in this section

    fmc_bookie = fmc_bookie.sort_values(['Name','ROI_(Horse_Wins)'],ascending = (False,False))
    fmc_bookie = fmc_bookie.reset_index(drop=True)

    bookie_drop_list = []
    for i in range(1, len(fmc_bookie)):
        if fmc_bookie.iloc[i]['Name'] == fmc_bookie.iloc[i-1]['Name']:
            bookie_drop_list.append(i)

    fmc_current_bookie = fmc_bookie.drop(bookie_drop_list)    
    print('NUM OF PROFITABLE BETS (B3,SK,PP): {} @ {}'.format(len(fmc_current_bookie),runtime))

    global fmc_current_bookie_email,fmc_fifth_email
    fmc_current_bookie_email = fmc_current_bookie[['ROI_(Horse_Wins)','ROI_(Horse_Loses)','Total Liability', 'Name','Location','Time','WIN Link','Oddschecker Link','PB Laystake','WB Laystake']]
    fmc_fifth_email = fmc_fifth[['Judge','SW','SP in SP','SP in AP','SP DP','Total Liability', 'Name','Location','Time','WIN Link','Oddschecker Link','PB Laystake','WB Laystake']]
# endregion
# region - EMAIL 
def email():       
    def automate(data): 
        mail_content = """\
        <html>
            <head></head>
            <body>
            {0}
            </body>
        </html>
        """.format(data.to_html())
        sender_address = '# ENTER EMAIL ADDRESS HERE'
        sender_pass = '# ENTER EMAIL PASSWORD HERE'
        receiver_address = str(values[5])
        #Setup the MIME
        message = MIMEMultipart()
        message['From'] = sender_address
        message['To'] = receiver_address
        message['Subject'] = 'EWS/Extra Place Data'    
        message.attach(MIMEText(mail_content, 'html'))
        session = smtplib.SMTP('smtp.gmail.com', 587) 
        session.starttls() 
        session.login(sender_address, sender_pass) 
        text = message.as_string()
        session.sendmail(sender_address, receiver_address, text)
        session.quit()

    if len(fmc_fifth_email) > 0:
        automate(fmc_fifth_email)
    if len(fmc_current_bookie_email) > 0:
        automate(fmc_current_bookie_email)
    exit()
# endregion
# region - END SCRIPT 
def end_script():
    closing_time = []
    global opening_time
    opening_time = []

    def closingtime():

        params = {
            'api_key': SCRAPINGBEE_API_KEY,
            'url': 'https://apieds.betfair.com/api/eds/meeting-races/v4?_ak=nzIFcwyWhrlwYMrh&countriesGroup=%5B%5B%22GB%22,'\
        '%22IE%22%5D%5D&countriesList=%5B%22GB%22,%22IE%22%5D&eventTypeId=7&marketStartingAfter={}-{}-{}T00:00:00.000Z&marketStarting'\
        'Before={}-{}-{}T23:59:59.999Z'.format(today.year,month,day,today.year,month,day)
        }

        content = str(requests.get(endpoint, params=params).content)

        starttime_list, starttime_list_x = ([] for i in range(2))
        needle2 = "startTime"
        for i, _ in enumerate(content):
            if content[i:i + len(needle2)] == needle2:
                starttime_list.append(content[i+23:i+28])
        for time in starttime_list:
            x = int(time.replace(':',''))
            starttime_list_x.append(x)

        max_value = max(starttime_list_x)
        min_value = min(starttime_list_x)
        max_index = starttime_list_x.index(max_value)
        closing_time.append(starttime_list[max_index])
        min_index = starttime_list_x.index(min_value)
        opening_time.append(starttime_list[min_index])

    closingtime()

    def exit_script():
        sys.exit()

    if int(runtime.replace(':','')) > int(closing_time[0].replace(':','')):
        if today.minute <= 9:
            close_time = ('%s:0%s' % (today.hour,today.minute))
            close_time.append(closing_time)
        else: # if today.minute >= 10:    
            close_time = ('%s:%s' % (today.hour,today.minute))
            close_time.append(closing_time)

    schedule.every().day.at(str(closing_time[-1])).do(exit_script)
# endregion
# region - FMC PROCESS
def fmc_process():
    run_time()
    end_script()
    sleep_time()
    racestoday()
    betfair()
    oddschecker_1()
    oddschecker_2()
    fmc_1()
    fmcfifth()
    fmc_2()
    email()
# endregion
# region - SLEEP and RUNTIME
def run_time():
    global today, month, day, year, runtime
    today = datetime.datetime.today()
    month = '%02d' % today.month
    day = '%02d' % today.day
    year = datetime.date.today().year
    if today.minute <= 9:
        runtime = ('%s:0%s' % (today.hour,today.minute))
    else:
        runtime = ('%s:%s' % (today.hour,today.minute))

def sleep_time():
    today_open = today.replace(hour=(int(opening_time[0][0:2])-1), minute=int(opening_time[0][3:5]))
    global sleep
    if today < today_open: 
        sleep = int((today_open-today).total_seconds())
    else:
        sleep = int(values[12])
# endregion
# region - RUN SCRIPT
while True: 
    try:
        fmc_process()
        time.sleep(sleep)

    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        traceback.print_exception(exc_type, exc_obj, exc_tb)
        time.sleep(sleep)
# endregion
