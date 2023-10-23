import pandas as pd
import datetime
import re
import numpy as np
import sys
sys.path.insert(1, 'Module')
from __main__ import r_LG, r_PV_F, r_PV_DS, r_PV_DOW, dach_süd, dach_ostwest, fassade_süd, fassade_west, fassade_nord, fassade_ost, Feiertage, r_PV_CP, carport

print("Lastgang lesen und PV Eigenverbrauch berechnen")

#create dataframe with the defined excel
r_LG['Time'] = r_LG['Time'].dt.floor('Min')
data_S = r_LG
del data_S["Einheit"]

#PV Anlage einlesen
if dach_süd == 1:
    data_S = data_S.join(r_PV_DS)
    del data_S["Datum"]
    del data_S["Photovoltaik  1: Energieproduktion AC (aus PV) [W] (Qinvpv)"]
if dach_süd == 0:
    data_S["dach_süd"] = 0
if dach_ostwest == 1:
    data_S = data_S.join(r_PV_DOW)
    del data_S["Datum"]
    del data_S["dach_west"]
    del data_S["dach_ost"]
    del data_S["west skalierung"]
    del data_S["ost skalierung"]
    del data_S["Photovoltaik  west: Energieproduktion AC (aus PV) [W] (Qinvpv)"]
    del data_S["Photovoltaik ost: Energieproduktion AC (aus PV) [W] (Qinvpv)"]
if dach_ostwest == 0:
    data_S["dach_ostwest"] = 0
if carport == 1:
    data_S = data_S.join(r_PV_CP)
    del data_S["Datum"]
    del data_S["dach_west"]
    del data_S["dach_ost"]
    del data_S["west skalierung"]
    del data_S["ost skalierung"]
    del data_S["Photovoltaik  west: Energieproduktion AC (aus PV) [W] (Qinvpv)"]
    del data_S["Photovoltaik ost: Energieproduktion AC (aus PV) [W] (Qinvpv)"]
if carport == 0:
    data_S["carport"] = 0
if fassade_süd + fassade_west + fassade_nord + fassade_ost >= 1:
    data_S = data_S.join(r_PV_F)
    del data_S["Datum"]
    del data_S["Photovoltaik süd: Energieproduktion AC (aus PV) [W] (Qinvpv)"]
    del data_S["Photovoltaik  west: Energieproduktion AC (aus PV) [W] (Qinvpv)"]
    del data_S["Photovoltaik nord: Energieproduktion AC (aus PV) [W] (Qinvpv)"]
    del data_S["Photovoltaik ost: Energieproduktion AC (aus PV) [W] (Qinvpv)"]
if fassade_süd == 0:
    data_S["fassade_süd"] = 0
if fassade_west == 0:
    data_S["fassade_west"] = 0
if fassade_nord == 0:
    data_S["fassade_nord"] = 0
if fassade_ost == 0:
    data_S["fassade_ost"] = 0

#calculate Lastgang and Rueckspeisung with PV-Eigenverbrauch
data_S["kW_Last"] = None
data_S['kW_Last'] = data_S['kW_netzbezug']
data_S["kW_PV_S"] = None
data_S["kW_RS_S"] = None
a = data_S.columns.get_loc('kW_Last')
b1 = data_S.columns.get_loc('dach_süd')
b2 = data_S.columns.get_loc('dach_ostwest')
b3 = data_S.columns.get_loc('fassade_süd')
b4 = data_S.columns.get_loc('fassade_west')
b5 = data_S.columns.get_loc('fassade_nord')
b6 = data_S.columns.get_loc('fassade_ost')
b7 = data_S.columns.get_loc('carport')
c = data_S.columns.get_loc('kW_PV_S')
d = data_S.columns.get_loc('kW_RS_S')
for row in range(0, len(data_S)):
    calc = round(data_S.iat[row, a] - (data_S.iat[row, b1] + data_S.iat[row, b2] + data_S.iat[row, b3] + data_S.iat[row, b4] + data_S.iat[row, b5] + data_S.iat[row, b6] + data_S.iat[row, b7]))
    if calc <= 0:
        lastP = 0
        rueck = calc*(-1)
    if calc > 0:
        lastP = calc*1
        rueck = 0
    data_S.iat[row, c] = lastP
    data_S.iat[row, d] = rueck

# Create column for Date
data_S['new_Date'] = None
index_set = data_S.columns.get_loc('Time')
index_date = data_S.columns.get_loc('new_Date')
for row in range(0, len(data_S)):
    Date: str = datetime.datetime.date(data_S.iat[row, index_set])
    data_S.iat[row, index_date] = Date

# array with only dates from Format 01.01.2019 00:15 for title:
date = [data_S.loc[:, "new_Date"][x:x + 96] for x in range(0, len(data_S), 96)]
def getUniqueItems(iterable):
    seen = set()
    result = []
    for item in iterable:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result
date_title = data_S.loc[:, "new_Date"]
title = getUniqueItems(date_title)
date_titles = []
for i in title:
    date_titles.append(i.strftime('%Y-%m-%d'))

data_S['Uhrzeit'] = None
index_set = data_S.columns.get_loc('Time')
index_zeit = data_S.columns.get_loc('Uhrzeit')
for row in range(0, len(data_S)):
    zeit: str = datetime.datetime.time(data_S.iat[row, index_set])
    data_S.iat[row, index_zeit] = zeit

data_S['month-day'] = None
index_set = data_S.columns.get_loc('new_Date')
index_moda = data_S.columns.get_loc('month-day')
for row in range(0, len(data_S)):
    moda: str = data_S.iat[row, index_set].strftime('%m-%d')
    data_S.iat[row, index_moda] = moda

data_S['Weekday'] = None
index_set = data_S.columns.get_loc('new_Date')
index_day = data_S.columns.get_loc('Weekday')
for row in range(0, len(data_S)):
    day: str = datetime.datetime.weekday(data_S.iat[row, index_set])
    data_S.iat[row, index_day] = day
    if any(str(data_S.iat[row, index_moda]) in s for s in Feiertage):
        data_S.iat[row, index_day] = 7

data_S['month'] = None
index_set = data_S.columns.get_loc('new_Date')
index_month = data_S.columns.get_loc('month')
for row in range(0, len(data_S)):
    month: str = data_S.iat[row, index_set].month
    data_S.iat[row, index_month] = month