# -*- coding: utf-8 -*-
"""
Created on Sat Mar 13 20:11:34 2021

@author: nou
"""

# you MUST install these packages seperately with pip install

import pandas as pd
import requests
import xlwings as xw
import datetime
import time


# READ ME

# 4 hour rolling average/min/max is calculated by storing the results of each minute's pull
# it will calculate these based on either all data it has or the last 4*60 updates, whichever is less
# therefore the program will become more accurate as it continues to run and store more pricing updates


########## about the wiki API

# Also read this: https://oldschool.runescape.wiki/w/RuneScape:Real-time_Prices

# set a helpful user agent, this info will be passed on to the wiki's servers
# read this section before setting a user agent https://oldschool.runescape.wiki/w/RuneScape:Real-time_Prices#Please_set_a_descriptive_User-Agent!

user_agent = 'margin analyzer @ nou#0999'

################################### SETUP - program needs to know excel tab names #############

# everything here needs to match the tab or file name of an excel file in the same folder as this script
# nothing will work unless this information matches
# NOTE: this will override anything in the tabs listed below, leave them blank for the program to use

current_price_tab = 'Prices'
previous_price_tab = 'Prices Previous'

fifteen_minute_low = '15 minute low'
fifteen_minute_high = '15 minute high'

four_hour_high = '4 hour high'
four_hour_low = '4 hour low'

book_name = 'OSRS GE.xlsx'

buy_limit_data_tab = 'buy limit data'

###################################################################################################


# initialize various things
try:
    cumulative_data
except NameError:
    cumulative_data = {'high': None,
                       'low': None}

failcount = 0
current_hour = datetime.datetime.now().hour

def write(sn, data, attempt_no=0):
    try:
        xw.Book(book_name).sheets[sn].range('A1').value = data
    except TypeError:
        xw.Book(book_name).sheets[sn].range('A1').value = data.to_frame()
    except:
        print("failed to write: " + sn)

        # yes this really helps, no I don't know why it randomly fails to write
        if attempt_no <= 5:
            write(sn, data, attempt_no+1)


def get_buy_limits():
    item_ids = requests.get('https://raw.githubusercontent.com/osrsbox/osrsbox-db/master/data/items/items-buylimits.json').json()
    write(buy_limit_data_tab, item_ids)
    print('updated buy limit lookup')

def store_data(cumulative_data, data):
    print("storing data")

    try:
        current_time = len(cumulative_data['high'].columns)
        print(current_time)
        low_data = data['low'].to_frame().rename(columns={'low': current_time})
        high_data = data['high'].to_frame().rename(columns={'high': current_time})
        cumulative_data['high'] = cumulative_data['high'].merge(
                high_data, left_index=True, right_index=True, how='outer')

        cumulative_data['low'] = cumulative_data['low'].merge(
                low_data, left_index=True, right_index=True, how='outer')
    except AttributeError:
        print('join failed, dumping cumulative database and starting over, this should happen when the program first runs (EAFP amiright)')
        cumulative_data['high'] = data['high'].to_frame().rename(columns={'high': 0})
        cumulative_data['low'] = data['low'].to_frame().rename(columns={'low': 0})

    return cumulative_data


def generate_min_max(data):
    n_cols = len(data['high'].columns)

    try:
        data['high'] = data['high'].to_frame()
        data['low'] = data['low'].to_frame()
    except AttributeError:
        pass

    # double transpose because iT's mOrE pYtHoNic and I can't get axis=1 to work
    do_calcs = lambda d, n: d[d.columns[-min(n_cols, n):]].T.aggregate(["min", "max", "mean"]).T

    if n_cols >= 15:
        write(fifteen_minute_high, do_calcs(data['high'], 15))
        write(fifteen_minute_low, do_calcs(data['low'], 15))

    # average since start or 4 hours
    write(four_hour_high, do_calcs(data['high'], 4*60))
    write(four_hour_low, do_calcs(data['low'], 4*60))


def get_data():
    raw_page_text = requests.get(
            'https://prices.runescape.wiki/api/v1/osrs/latest', headers={
                    'User-Agent': user_agent}).text
        
    # convert to a format read_json likes (2d data structure)
    short_page_text = raw_page_text[8:-1]

    table = pd.read_json(short_page_text)
    return table.T


def paste_data(data, previous_data):
    write(current_price_tab, data)
    write(previous_price_tab, previous_data)


def update_prices(previous_data):
    data = get_data()
    paste_data(data, previous_data)
    return data


get_buy_limits()
while True:
    if datetime.datetime.now().second == 1:
        print(datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        try:
            previous_data = update_prices(previous_data)
        except NameError:
            previous_data = get_data()
            continue
        except:
            print("failed")
            failcount += 1
            if failcount == 10:
                break
            continue
        cumulative_data = store_data(cumulative_data, previous_data)
        generate_min_max(cumulative_data)
        time.sleep(5)
    
    if current_hour != datetime.datetime.now().hour:
        failcount = 0
        current_hour = datetime.datetime.now().hour
    
    time.sleep(0.5)

# rip v1
#xw.Book('OSRS GE.xlsx').sheets['Prices'].range('A1').value = pd.read_json(
#        requests.get('https://prices.runescape.wiki/api/v1/osrs/latest', headers={
#                'User-Agent': 'Margin Analyzer'}).text).T
