import pandas as pd
import numpy as np
from itertools import compress

"Programm für die Analyse des Peak-Shaving Potenzials pro Speichergrösse. Gibt Auskunft über die maximale Reduktion ohne den Speicher zwischendurch zu laden."
def analyse_peaks(LG_A,Speicher,Faktor_Grenze,date_titles,title,ladeverlust):
    print("Analyse für monatliche Leistungs Limite Peak-Shaving")

    # list with kW, ordered per day from highest value to lowest
    sort_lst_kW = []
    kW_seq = [LG_A.loc[:, "ANALYSE_SPITZE"][x:x + 96] for x in range(0, len(LG_A), 96)]
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
                if sum_kWh >= S*(1-ladeverlust):
                    if counter < 2:
                        new_limit = np.max(sort_lst_kW[row_number]) - S*(1-ladeverlust)
                        limit_per_day.append(new_limit)
                        break
                    if counter >= 2:
                        new_limit = np.max(sort_lst_kW[row_number]) - sum_limit - (S*(1-ladeverlust)-(sum_kWh-(i*(-1) * counter)/4))*4/counter
                        if sum_kWh == S*(1-ladeverlust):
                            sum_limit = i * (-1) + sum_limit
                            new_limit = np.max(sort_lst_kW[row_number]) - sum_limit
                        limit_per_day.append(new_limit)
                        break
                sum_limit = i*(-1) + sum_limit
                counter = counter + 1
            if sum_kWh < S*ladeverlust:
                new_limit = np.max(sort_lst_kW[row_number]) - sum_limit - ((S*ladeverlust - sum_kWh) / counter * 4)
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

    # new monthly limit caused by faktor Grenze
    new_diff_for_max_limit = []
    for row in load_red:
        calc1 = []
        for i in row:
            calc1.append(i*Faktor_Grenze)
        new_diff_for_max_limit.append(calc1)
    new_monthly_limit = []
    counter1 = 0
    for row in limit_month_all_Storage:
        counter2 = 0
        calc1 = []
        for i in row:
            calc1.append(i + new_diff_for_max_limit[counter1][counter2])
            counter2 = counter2 + 1
        counter1 = counter1 + 1
        new_monthly_limit.append(calc1)
    new_monthly_limit = pd.DataFrame.from_dict(dict(zip(Speicher, new_monthly_limit)), orient='index').transpose()
    new_monthly_limit.insert(0, "month_index", [1,2,3,4,5,6,7,8,9,10,11,12], True)

    LG_S = pd.merge_ordered(LG_A, new_monthly_limit, left_on="month", right_on="month_index", how="left", fill_method="ffill")
    del LG_S["month_index"]
    return LG_S, new_monthly_limit
