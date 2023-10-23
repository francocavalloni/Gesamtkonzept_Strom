import pandas as pd
import numpy as np
from itertools import compress
from read_1 import data_S, date_titles, title
from Sim_V2 import LG_S
from __main__ import output_excel_A, spalte_analyse, Speicher, Faktor_Grenze
import sys
sys.path.insert(1, 'Module')

print("Analyse für monatliche Spitze")

LG_A = LG_S

# list with kW, ordered per day from highest value to lowest
sort_lst_kW = []
kW_seq = [LG_A.loc[:, spalte_analyse][x:x + 96] for x in range(0, len(LG_A), 96)]
for row in np.array(kW_seq):
    col_kW = sorted(row, key=np.max, reverse=True)
    sort_lst_kW.append(col_kW)
# arrays with the difference in kW to get to the next lower kW-peak
for row in sort_lst_kW:
    diff = np.diff(sort_lst_kW)
#TODO Anpassen weil nicht der maximalen reduktion entspricht!
# berechne limit pro tag in kW pro speichergrösse
limit_Speicher = []
limit_month_all_Storage = []
limit_per_day = []
for S in Speicher:
    row_number = 0
    for row in diff:
        counter = 1
        sum_kWh = 0
        sum_limit = 0
        for i in row:
            sum_kWh = (i*(-1) * counter)/4 + sum_kWh
            if sum_kWh >= S:
                if counter < 2:
                    new_limit = np.max(sort_lst_kW[row_number]) - sum_limit
                    limit_per_day.append(new_limit)
                    break
                if counter > 2:
                    new_limit = np.max(sort_lst_kW[row_number]) - sum_limit
                    limit_per_day.append(new_limit)
                    break
            sum_limit = i*(-1) + sum_limit
            counter = counter + 1
        if sum_kWh < S:
            new_limit = np.max(sort_lst_kW[row_number]) - sum_limit - ((S - sum_kWh) / counter * 4)
            limit_per_day.append(new_limit)
        row_number = row_number + 1
    lim_day = pd.DataFrame(limit_per_day)
    limit_per_day = lim_day[0].to_list()
    limit_Speicher.append(limit_per_day)
    # slice kW limit for months
    mon_ind = pd.DataFrame(index=pd.date_range(start=date_titles[0], end=date_titles[-1]))
    lst_kW_monthly = []
    for i in range(1, 13):
        lst_kW_monthly.append(list(compress(limit_per_day, mon_ind.index.month == i)))
    # max limit kW for every month
    Limit_monthly = []
    for column in lst_kW_monthly:
        max = np.max(column)
        Limit_monthly.append(max)
    limit_month_all_Storage.append(Limit_monthly)
    limit_per_day = []

month_titles = pd.date_range(title[0],title[-1], freq='MS').strftime("%b-%Y").tolist()

#berechnung der lastreduktion
load_red = []
for row in limit_month_all_Storage:
    counter = 0
    calc1 = []
    for i in row:
        calc1.append(limit_month_all_Storage[0][counter]-i)
        counter = counter + 1
    load_red.append(calc1)

#prepare dict to write in excel
daily_limit_all_Storage = pd.DataFrame.from_dict(dict(zip(Speicher, limit_Speicher)), orient='index').transpose()
daily_limit_all_Storage.insert(0, "Date", date_titles, True)
monthly_limit_all_Storage = pd.DataFrame.from_dict(dict(zip(Speicher, limit_month_all_Storage)), orient='index').transpose()
monthly_limit_all_Storage.insert(0, "month", month_titles, True)
monthly_load_red_all_Storage = pd.DataFrame.from_dict(dict(zip(Speicher, load_red)), orient='index').transpose()
monthly_load_red_all_Storage.insert(0, "month", month_titles, True)

# new monthly limit
new_load_red = []
for row in load_red:
    calc1 = []
    for i in row:
        calc1.append(i-i*(1-Faktor_Grenze))
    new_load_red.append(calc1)
new_monthly_limit = []
counter1 = 0
for row in limit_month_all_Storage:
    counter2 = 0
    calc1 = []
    for i in row:
        calc1.append(i + new_load_red[counter1][counter2])
        counter2 = counter2 + 1
    counter1 = counter1 + 1
    new_monthly_limit.append(calc1)
new_monthly_limit = pd.DataFrame.from_dict(dict(zip(Speicher, new_monthly_limit)), orient='index').transpose()
new_monthly_limit.insert(0, "month_index", [1,2,3,4,5,6,7,8,9,10,11,12], True)

LG_S = pd.merge_ordered(LG_S, new_monthly_limit, left_on="month", right_on="month_index", how="left", fill_method="ffill")
del LG_S["month_index"]
