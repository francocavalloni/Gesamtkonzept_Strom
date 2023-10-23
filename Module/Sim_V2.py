import pandas as pd
import datetime
import sys
sys.path.insert(1, 'Module')
from read_1 import date_titles, title, data_S
from __main__ import LG_name, Sim_LG, öl_verbrauch, PV_neu, Faktor_Grenze, Faktor_laden, Eigenverbrauch_grenze, output_excel_S, Speicher, Eigenverbrauch, \
    leistungsspitzenpreis, hochtarif, niedertarif, rueckspeisungstarif, öl_tarif,\
    JahreskostenPV_var, A_BS, Faktor_vergünstigung, Capex_PV, jährliche_Ausgaben_PV, EIV, \
    ELKW_Sim, elkw1_S, elkw2_S, elkw3_S, Verbrauch, ladeleistungDC, Kapazität_Wechselbatterie, ELKW_SOC_limit_Ruhezeit, schnellladung, reserve_akku , tarif_schnellladen, Grenze_SOC_Raststätte_laden, laden_30min_raststätte, JahreskostenELKW_var, Jahreskosten_dc_ladestation, \
    EPKW_Sim, ac_ladestation_inv, ac_ladestation_unterhalt, durchsch_kapazität, durchsch_Reichweite, ladeleistung_stationen, epkw_ladekapazitaet, epkw_wochenstart_kWh, JahreskostenLKW_var, N_Ladestationen, Jahreskosten_ac_ladestation,\
    WELKW_Sim, Kapazität_Wechselbatterie, welkw1_S, welkw2_S, welkw3_S, SOC_fuer_wechsel, Kosten_Wechselbatterie, JahreskostenWELKW_var
if ELKW_Sim >=1 or WELKW_Sim>=1:
    from read_M_1 import data_S, elkw1_summe_break_diff, elkw2_summe_break_diff, elkw3_summe_break_diff

print("Simulation Running")

# Dataframes Resultate Jahreswerte BILANZ TOTAL
totname1 = "Netzbezug total [kWh/a]"
totname2 = "Rückspeisung total [kWh/a]"
totname3 = "Lastspitzen total [kW/a]"
totname4 = "Auswärtsladen elkw + Batteriewechsel bei Wechselsystem total [kWh/a]"
totname5 = "Eigenverbrauch total [kWh/a]"
totname6 = "Lastspitzenreduktion total [kW/a]"
totname7 = "Kosten Netzbezug [CHF/a]"
totname8 = "Kosten Lastspitzen [CHF/a]"
totname9 = "Kosten Auswärtsladen + Batteriewechsel bei Wechselsystem [CHF/a]"
totname10 = "Kosten Strom [CHF/a]"
totname11 = "Einnahmen Rückspeisung [CHF/a]"
totname12 = "Kosten Thermoöl [CHF/a]"
totname13 = "Jahreskosten Investition und Unterhalt [CHF/a]"
totname14 = "Kosten total Variante [CHF/a]"
names_BILANZ_TOTAL = ["jährliche Lastreduktion [kW/a]", "jährliche Eigenverbrauchserhöhung [kWh/a]", "jährliche Einsparung_PS [CHF/a]", "jährliche Einsparung_EV [CHF/a]", "jährlicher Gewinn Rückspeisung [CHF/a]", "Jahreskosten [CHF/a]", "Businesscase PV und BS [CHF/a]", ]
names_BILANZ_TOTAL = [totname1, totname2, totname3, totname4, totname5, totname6, totname7, totname8, totname9, totname10, totname11, totname12, totname13, totname14]
BILANZ_TOTAL = pd.DataFrame(index=range(len(names_BILANZ_TOTAL)), columns=[output_excel_S])
BILANZ_TOTAL[output_excel_S] = names_BILANZ_TOTAL

LG_S = data_S
LG_S["ANALYSE_SPITZE"] = LG_S["kW_PV_S"]
columns = ['Time']
last_best = LG_S.columns.get_loc("kW_Last")
last = LG_S.columns.get_loc('kW_PV_S')
rueck = LG_S.columns.get_loc('kW_RS_S')
weekday = LG_S.columns.get_loc('Weekday')
month = LG_S.columns.get_loc('month')
zeit = LG_S.columns.get_loc('Uhrzeit')
date = LG_S.columns.get_loc('new_Date')

month_titles = pd.date_range(title[0],title[-1], freq='MS').strftime("%b-%Y").tolist()

counterLR = 0

if ELKW_Sim >= 1:
    from Analyse_V1 import LG_S, new_monthly_limit
    #Set Dataframes
    names_Bilanz_ELKW = ["Speicherkapazität [kWh]", "Verbrauch [kWh/km]", "Energiebezug Speicher [kWh/a]", "Netzbezug [kWh/a]", "Eigenverbrauchserhöhung [kWh/a]",
                         "Auswärtsladen [kWh/a]",  "Mittagspause Verlängert um [min]", "Häufigkeit Pause an Raststätte 30min laden",
                        "Kosten Netzbezug [CHF/a]","Kosten Laden Auswärts [CHF/a]","Kosten Eigenverbauch [CHF/a]",
                         "Summe Energiekosten [CHF/a]", "Jahreskosten Inv + Unt [CHF/a]", "Jahreskosten Ladestation DC Inv + Unt", "Kosten total ELKW [CHF/a]"]
    BILANZ_ELKW = pd.DataFrame(index=range(len(names_Bilanz_ELKW)), columns=[output_excel_S])
    BILANZ_ELKW[output_excel_S] = names_Bilanz_ELKW
    BILANZ_ELKW["ELKW1"] = None
    BILANZ_ELKW["ELKW2"] = None
    BILANZ_ELKW["ELKW3"] = None
    BILANZ_ELKW["Summe 3 ELKWs"] = None
    INFO_ELKW = pd.DataFrame(index=range(len(LG_S)), columns=columns)
    INFO_ELKW["Time"] = LG_S["Time"]
    INFO_ELKW["Weekday"] = LG_S["Weekday"]
    INFO_ELKW["ELKW1_Warnung"] = None
    INFO_ELKW["ELKW1_SOC"] = None
    INFO_ELKW["ELKW2_Warnung"] = None
    INFO_ELKW["ELKW2_SOC"] = None
    INFO_ELKW["ELKW3_Warnung"] = None
    INFO_ELKW["ELKW3_SOC"] = None
    for elkw in range(1,ELKW_Sim+1):
        counter_title = elkw
        if elkw ==1:
            elkw_S = elkw1_S
            elkw_summe_break_diff = elkw1_summe_break_diff
            INFO_ELKW["Last_ELKW1"] = LG_S["kW_PV_S"]
            INFO_ELKW["Ruecksp_ELKW1"] = LG_S["kW_RS_S"]
        if elkw ==2:
            elkw_S = elkw2_S
            elkw_summe_break_diff = elkw2_summe_break_diff
            INFO_ELKW["Last_ELKW2"] = INFO_ELKW["Last_ELKW1"]
            INFO_ELKW["Ruecksp_ELKW2"] = INFO_ELKW["Ruecksp_ELKW1"]
        if elkw ==3:
            elkw_S = elkw3_S
            elkw_summe_break_diff = elkw3_summe_break_diff
            INFO_ELKW["Last_ELKW3"] = INFO_ELKW["Last_ELKW2"]
            INFO_ELKW["Ruecksp_ELKW3"] = INFO_ELKW["Ruecksp_ELKW2"]
        elkw_last = INFO_ELKW.columns.get_loc('Last_ELKW' + str(counter_title))
        elkw_rueck = INFO_ELKW.columns.get_loc('Ruecksp_ELKW' + str(counter_title))
        elkw_warnung = INFO_ELKW.columns.get_loc('ELKW' + str(counter_title) + '_Warnung')
        elkw_soc = INFO_ELKW.columns.get_loc('ELKW' + str(counter_title) + '_SOC')
        elkw_Energiebezug = 0
        elkw_Eigenverbrauch = 0
        elkw_Netzbezug = 0
        elkw_SOC = 1
        elkw_GES_verbauch = 0
        elkw_summe_preis = 0
        elkw_netzbezugauswaerts = 0
        elkw_counter_rast = 0
        Raststätte_verzögerung = 0
        for row in range(0, len(INFO_ELKW)):
            last_S = INFO_ELKW.iat[row, elkw_last]
            ruecksp = INFO_ELKW.iat[row, elkw_rueck]
            date_S = LG_S.iat[row, date]
            zeit_S = LG_S.iat[row, zeit]
            weekday_S = LG_S.iat[row, weekday]
            month_S = LG_S.iat[row, month]
            status_elkw = LG_S.iat[row, LG_S.columns.get_loc(str(counter_title)+' status')]
            distance_elkw = LG_S.iat[row, LG_S.columns.get_loc(str(counter_title)+' driving_distance')]
            beladen_elkw = LG_S.iat[row, LG_S.columns.get_loc(str(counter_title)+' Anzahl beladen Lupfig')]
            totdistance_elkw = LG_S.iat[row, LG_S.columns.get_loc(str(counter_title)+' total_distance')]
            ps_grenze = LG_S.iat[row, LG_S.columns.get_loc(0)]
            elkw_verbrauch = 0
            elkw_ladeleistung = 0
            elkw_ladeleistung_r = 0
            elkw_ladeleistung_n = 0
            elkw_ladeleistung_a = 0
            if ELKW_Sim >= 1:
                #### unterwegs
                if status_elkw == 1:
                    elkw_verbrauch = distance_elkw * Verbrauch
                    elkw_SOC = elkw_SOC - (elkw_verbrauch / elkw_S)
                    elkw_GES_verbauch = elkw_GES_verbauch + elkw_verbrauch
                    if elkw_SOC <= 0:
                        INFO_ELKW.iat[row, elkw_warnung] = "Batterie leer"
                #### laden mit PV-Überschuss # status: 0 (lupfig ausserhalb betrieb),-1 (beladen_lupfig), -2 (mittagspause), -3 (beladen auswärts), -4 (beladen taracelll burnhaupt)
                if elkw_SOC < 1 and status_elkw == 0 or elkw_SOC < 1 and status_elkw == (-2) and beladen_elkw > 1\
                        or elkw_SOC < 1 and status_elkw == (-1):
                    if ruecksp >= ladeleistungDC:
                        elkw_ladeleistung = ladeleistungDC
                        elkw_SOC_t = elkw_SOC + elkw_ladeleistung / 4 / elkw_S
                        if elkw_SOC_t > 1:
                            elkw_ladeleistung = (1 - elkw_SOC) * 4 * elkw_S
                        elkw_SOC = elkw_SOC + elkw_ladeleistung / 4 / elkw_S
                    if ruecksp < ladeleistungDC and ruecksp > 0:
                        elkw_ladeleistung = ruecksp
                        elkw_SOC_t = elkw_SOC + elkw_ladeleistung / 4 / elkw_S
                        if elkw_SOC_t > 1:
                            elkw_ladeleistung = (1 - elkw_SOC) * 4 * elkw_S
                        elkw_SOC = elkw_SOC + elkw_ladeleistung / 4 / elkw_S
                    elkw_ladeleistung_r = elkw_ladeleistung
                    ruecksp = ruecksp - elkw_ladeleistung
                    elkw_Energiebezug = elkw_Energiebezug + elkw_ladeleistung / 4
                    elkw_Eigenverbrauch = elkw_Eigenverbrauch + elkw_ladeleistung / 4
                #### laden mit netz
                if elkw_SOC < ELKW_SOC_limit_Ruhezeit and elkw_ladeleistung_r < ladeleistungDC and status_elkw == 0 \
                    or totdistance_elkw > (1 - reserve_akku) * elkw_S * elkw_SOC / Verbrauch and elkw_SOC < 1 and status_elkw == 0 and elkw_ladeleistung_r < ladeleistungDC and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") < datetime.datetime.strptime("12:00:00", "%H:%M:%S") \
                    or totdistance_elkw / 2 > (1 - reserve_akku) * elkw_S * elkw_SOC / Verbrauch and elkw_SOC < 1 and status_elkw == -2 and elkw_ladeleistung_r < ladeleistungDC and beladen_elkw > 1\
                    or totdistance_elkw > (1 - reserve_akku) * elkw_S * elkw_SOC / Verbrauch and elkw_SOC < 1 and status_elkw == -1 and elkw_ladeleistung_r < ladeleistungDC:
                    ladeleistung_netz = ladeleistungDC
                    if datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") < datetime.datetime.strptime("04:00:00", "%H:%M:%S") and last_S + ladeleistungDC*(4-elkw) > ps_grenze:
                        ladeleistung_netz = (ps_grenze-last_S)/elkw #kW
                    elkw_ladeleistung = ladeleistung_netz - elkw_ladeleistung_r
                    elkw_SOC_t = elkw_SOC + elkw_ladeleistung / 4 / elkw_S
                    if elkw_SOC_t > 1:
                        elkw_ladeleistung = (1 - elkw_SOC) * 4 * elkw_S
                    last_S = last_S + elkw_ladeleistung
                    elkw_ladeleistung_n = elkw_ladeleistung
                    elkw_SOC = elkw_SOC + elkw_ladeleistung / 4 / elkw_S
                    elkw_Energiebezug = elkw_Energiebezug + elkw_ladeleistung / 4
                    elkw_Netzbezug = elkw_Netzbezug + elkw_ladeleistung / 4
                #### laden auswärts, wenn soc zu tief für rückweg, mit eingerechneter reserve von 10%, mittagspause
                if totdistance_elkw / 2 > (1 - reserve_akku) * elkw_S * elkw_SOC / Verbrauch and elkw_SOC < 1 and status_elkw == (-2) and beladen_elkw == 1 and elkw_ladeleistung == 0:
                    elkw_ladeleistung = schnellladung
                    elkw_SOC_t = elkw_SOC + elkw_ladeleistung / 4 / elkw_S
                    if elkw_SOC_t > 1:
                        elkw_ladeleistung = (1 - elkw_SOC) * 4 * elkw_S
                    elkw_SOC = elkw_SOC + elkw_ladeleistung / 4 / elkw_S
                    ladeleistung_elkw1_a = elkw_ladeleistung
                    elkw_netzbezugauswaerts = elkw_netzbezugauswaerts + elkw_ladeleistung / 4
                    elkw_Energiebezug = elkw_Energiebezug + elkw_ladeleistung / 4
                if status_elkw == 1 and elkw_SOC < Grenze_SOC_Raststätte_laden and Raststätte_verzögerung == 0:
                    Raststätte_verzögerung = 2
                    target_date = date_S
                    target_time =  datetime.datetime.combine(date_S, zeit_S)
                    target_time += datetime.timedelta(minutes=15)
                    target_time = target_time.time()
                    mask = (LG_S['new_Date'] == target_date) & (LG_S['Uhrzeit'] >= target_time)
                    LG_S.loc[mask, str(counter_title)+' driving_distance'] = LG_S.loc[mask, str(counter_title)+' driving_distance'].shift(+2)
                    LG_S.loc[mask, str(counter_title)+' status'] = LG_S.loc[mask, str(counter_title)+' status'].shift(+2)
                    target_time = zeit_S
                if Raststätte_verzögerung > 0 and zeit_S > target_time:
                    elkw_SOC = elkw_SOC + laden_30min_raststätte / elkw_S /2
                    elkw_netzbezugauswaerts = elkw_netzbezugauswaerts + laden_30min_raststätte/2
                    elkw_Energiebezug = elkw_Energiebezug + laden_30min_raststätte/2
                    elkw_counter_rast = elkw_counter_rast + 1
                    Raststätte_verzögerung = Raststätte_verzögerung - 1
                INFO_ELKW.iat[row, elkw_last] = last_S
                INFO_ELKW.iat[row, elkw_rueck] = ruecksp
                INFO_ELKW.iat[row, elkw_soc] = elkw_SOC
                if (weekday_S < 5 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") > datetime.datetime.strptime("07:00:00", "%H:%M:%S")
                    and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("20:00:00","%H:%M:%S")) \
                    or (weekday_S == 5 and datetime.datetime.strptime(str(zeit_S),"%H:%M:%S") > datetime.datetime.strptime("07:00:00", "%H:%M:%S")
                    and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("13:00:00", "%H:%M:%S")):
                    preis = elkw_ladeleistung_n / 4 * hochtarif
                else:
                    preis = elkw_ladeleistung_n / 4 * niedertarif
                elkw_summe_preis = elkw_summe_preis + preis
                LG_S.iat[row, LG_S.columns.get_loc("ANALYSE_SPITZE")] = last_S
        BILANZ_ELKW.iat[0, counter_title] = elkw_S
        BILANZ_ELKW.iat[1, counter_title] = Verbrauch
        BILANZ_ELKW.iat[2, counter_title] = elkw_Energiebezug
        BILANZ_ELKW.iat[3, counter_title] = round(elkw_Netzbezug)
        BILANZ_ELKW.iat[4, counter_title] = round(elkw_Eigenverbrauch)
        BILANZ_ELKW.iat[5, counter_title] = round(elkw_netzbezugauswaerts)
        BILANZ_ELKW.iat[6, counter_title] = elkw_summe_break_diff
        BILANZ_ELKW.iat[7, counter_title] = elkw_counter_rast/2
        BILANZ_ELKW.iat[8, counter_title] = round(elkw_summe_preis)
        BILANZ_ELKW.iat[9, counter_title] = round(elkw_netzbezugauswaerts*tarif_schnellladen)
        BILANZ_ELKW.iat[10, counter_title] = round(elkw_Eigenverbrauch*rueckspeisungstarif)
        BILANZ_ELKW.iat[11, counter_title] = round(elkw_summe_preis + elkw_netzbezugauswaerts*tarif_schnellladen + elkw_Eigenverbrauch*rueckspeisungstarif)
        BILANZ_ELKW.iat[12, counter_title] = round(JahreskostenELKW_var)
        BILANZ_ELKW.iat[13, counter_title] = round(Jahreskosten_dc_ladestation)
        BILANZ_ELKW.iat[14, counter_title] = round(JahreskostenELKW_var + Jahreskosten_dc_ladestation + round(elkw_summe_preis + elkw_netzbezugauswaerts*tarif_schnellladen + elkw_Eigenverbrauch*rueckspeisungstarif))
        counter_title = counter_title + 1
    for row in range(2, len(BILANZ_ELKW)):
        BILANZ_ELKW.iat[row, 4] = BILANZ_ELKW.iat[row, 1] + BILANZ_ELKW.iat[row, 2] + BILANZ_ELKW.iat[row, 3]

if WELKW_Sim >= 1:
    from Analyse_V1 import LG_S, new_monthly_limit
    #Set Dataframes
    names_Bilanz_WELKW = ["Speicherkapazität [kWh]", "Verbrauch [kWh/km]", "Energiebezug Speicher [kWh/a]", "Netzbezug [kWh/a]", "Eigenverbrauchserhöhung [kWh/a]",
                         "Auswärtsladen [kWh/a]",  "Mittagspause Verlängert um [min]", "Häufigkeit Batteriewechsel an Station [5min]", "Laden durch Batteriewechsel [kWh/a]",
                        "Kosten Netzbezug [CHF/a]","Kosten Laden Auswärts [CHF/a]","Kosten Eigenverbauch [CHF/a]", "Kosten Batteriewechsel [CHF/a",
                         "Summe Energiekosten [CHF/a]", "Jahreskosten Inv + Unt [CHF/a]", "Jahreskosten Ladestation DC Inv + Unt" ,  "Kosten total ELKW [CHF/a]"]
    BILANZ_WELKW = pd.DataFrame(index=range(len(names_Bilanz_WELKW)), columns=[output_excel_S])
    BILANZ_WELKW[output_excel_S] = names_Bilanz_WELKW
    BILANZ_WELKW["WELKW1"] = None
    BILANZ_WELKW["WELKW2"] = None
    BILANZ_WELKW["WELKW3"] = None
    BILANZ_WELKW["Summe 3 WELKWs"] = None
    INFO_WELKW = pd.DataFrame(index=range(len(LG_S)), columns=columns)
    INFO_WELKW["Time"] = LG_S["Time"]
    INFO_WELKW["Weekday"] = LG_S["Weekday"]
    INFO_WELKW["WELKW1_Warnung"] = None
    INFO_WELKW["WELKW1_SOC"] = None
    INFO_WELKW["WELKW2_Warnung"] = None
    INFO_WELKW["WELKW2_SOC"] = None
    INFO_WELKW["WELKW3_Warnung"] = None
    INFO_WELKW["WELKW3_SOC"] = None
    for welkw in range(1, WELKW_Sim + 1):
        counter_title = welkw
        if welkw ==1:
            welkw_S = welkw1_S
            elkw_summe_break_diff = elkw1_summe_break_diff
            INFO_WELKW["Last_WELKW1"] = LG_S["kW_PV_S"]
            INFO_WELKW["Ruecksp_WELKW1"] = LG_S["kW_RS_S"]
        if welkw ==2:
            welkw_S = welkw2_S
            elkw_summe_break_diff = elkw2_summe_break_diff
            INFO_WELKW["Last_WELKW2"] = INFO_WELKW["Last_WELKW1"]
            INFO_WELKW["Ruecksp_WELKW2"] = INFO_WELKW["Ruecksp_WELKW1"]
        if welkw ==3:
            welkw_S = welkw3_S
            elkw_summe_break_diff = elkw3_summe_break_diff
            INFO_WELKW["Last_WELKW3"] = INFO_WELKW["Last_WELKW2"]
            INFO_WELKW["Ruecksp_WELKW3"] = INFO_WELKW["Ruecksp_WELKW2"]
        welkw_last = INFO_WELKW.columns.get_loc('Last_WELKW' + str(counter_title))
        welkw_rueck = INFO_WELKW.columns.get_loc('Ruecksp_WELKW' + str(counter_title))
        welkw_warnung = INFO_WELKW.columns.get_loc('WELKW' + str(counter_title) + '_Warnung')
        welkw_soc = INFO_WELKW.columns.get_loc('WELKW' + str(counter_title) + '_SOC')
        welkw_Energiebezug = 0
        welkw_Eigenverbrauch = 0
        welkw_Netzbezug = 0
        welkw_SOC = 1
        welkw_GES_verbauch = 0
        welkw_summe_preis = 0
        welkw_netzbezugauswaerts = 0
        welkw_counter_wechsel = 0
        for row in range(0, len(INFO_WELKW)):
            last_S = INFO_WELKW.iat[row, welkw_last]
            ruecksp = INFO_WELKW.iat[row, welkw_rueck]
            date_S = LG_S.iat[row, date]
            zeit_S = LG_S.iat[row, zeit]
            weekday_S = LG_S.iat[row, weekday]
            month_S = LG_S.iat[row, month]
            status_welkw = LG_S.iat[row, LG_S.columns.get_loc(str(counter_title) + ' status')]
            distance_welkw = LG_S.iat[row, LG_S.columns.get_loc(str(counter_title) + ' driving_distance')]
            beladen_welkw = LG_S.iat[row, LG_S.columns.get_loc(str(counter_title) + ' Anzahl beladen Lupfig')]
            totdistance_welkw = LG_S.iat[row, LG_S.columns.get_loc(str(counter_title) + ' total_distance')]
            ps_grenze = LG_S.iat[row, LG_S.columns.get_loc(0)]
            welkw_verbrauch = 0
            welkw_ladeleistung = 0
            welkw_ladeleistung_r = 0
            welkw_ladeleistung_n = 0
            welkw_ladeleistung_a = 0
            if WELKW_Sim >= 1:
                #### unterwegs
                if status_welkw == 1:
                    welkw_verbrauch = distance_welkw * Verbrauch
                    welkw_SOC = welkw_SOC - (welkw_verbrauch / welkw_S)
                    welkw_GES_verbauch = welkw_GES_verbauch + welkw_verbrauch
                    if welkw_SOC <= 0:
                        INFO_WELKW.iat[row, welkw_warnung] = "Batterie leer"
                #### laden mit PV-Überschuss # status: 0 (lupfig ausserhalb betrieb),-1 (beladen_lupfig), -2 (mittagspause), -3 (beladen auswärts), -4 (beladen taracelll burnhaupt)
                if welkw_SOC < 1 and status_welkw == 0 or welkw_SOC < 1 and status_welkw == (-2) and beladen_welkw > 1\
                        or welkw_SOC < 1 and status_welkw == (-1):
                    if ruecksp >= ladeleistungDC:
                        welkw_ladeleistung = ladeleistungDC
                        elkw_SOC_t = welkw_SOC + welkw_ladeleistung / 4 / welkw_S
                        if elkw_SOC_t > 1:
                            welkw_ladeleistung = (1 - welkw_SOC) * 4 * welkw_S
                        welkw_SOC = welkw_SOC + welkw_ladeleistung / 4 / welkw_S
                    if ruecksp < ladeleistungDC and ruecksp > 0:
                        welkw_ladeleistung = ruecksp
                        elkw_SOC_t = welkw_SOC + welkw_ladeleistung / 4 / welkw_S
                        if elkw_SOC_t > 1:
                            welkw_ladeleistung = (1 - welkw_SOC) * 4 * welkw_S
                        welkw_SOC = welkw_SOC + welkw_ladeleistung / 4 / welkw_S
                    welkw_ladeleistung_r = welkw_ladeleistung
                    ruecksp = ruecksp - welkw_ladeleistung
                    welkw_Energiebezug = welkw_Energiebezug + welkw_ladeleistung / 4
                    welkw_Eigenverbrauch = welkw_Eigenverbrauch + welkw_ladeleistung / 4
                #### laden mit netz
                if welkw_SOC < ELKW_SOC_limit_Ruhezeit and welkw_ladeleistung_r < ladeleistungDC and status_welkw == 0 \
                    or totdistance_welkw > (1 - reserve_akku) * welkw_S * welkw_SOC / Verbrauch and welkw_SOC < 1 and status_welkw == 0 and welkw_ladeleistung_r < ladeleistungDC and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") < datetime.datetime.strptime("12:00:00", "%H:%M:%S") \
                    or totdistance_welkw / 2 > (1 - reserve_akku) * welkw_S * welkw_SOC / Verbrauch and welkw_SOC < 1 and status_welkw == -2 and welkw_ladeleistung_r < ladeleistungDC and beladen_welkw > 1\
                    or totdistance_welkw > (1 - reserve_akku) * welkw_S * welkw_SOC / Verbrauch and welkw_SOC < 1 and status_welkw == -1:
                    ladeleistung_netz = ladeleistungDC
                    if datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") < datetime.datetime.strptime("04:00:00","%H:%M:%S") and last_S + ladeleistungDC * (4 - welkw) > ps_grenze:
                        ladeleistung_netz = (ps_grenze - last_S) / welkw  # kW
                    welkw_ladeleistung = ladeleistung_netz - welkw_ladeleistung_r
                    elkw_SOC_t = welkw_SOC + welkw_ladeleistung / 4 / welkw_S
                    if elkw_SOC_t > 1:
                        welkw_ladeleistung = (1 - welkw_SOC) * 4 * welkw_S
                    last_S = last_S + welkw_ladeleistung
                    welkw_ladeleistung_n = welkw_ladeleistung
                    welkw_SOC = welkw_SOC + welkw_ladeleistung / 4 / welkw_S
                    welkw_Energiebezug = welkw_Energiebezug + welkw_ladeleistung / 4
                    welkw_Netzbezug = welkw_Netzbezug + welkw_ladeleistung / 4
                #### laden auswärts, wenn soc zu tief für rückweg, mit eingerechneter reserve von 10%, mittagspause, zb. burnhaupt/ecublens
                if totdistance_welkw / 2 > (1 - reserve_akku) * welkw_S * welkw_SOC / Verbrauch and welkw_SOC < 1 and status_welkw == (-2) and beladen_welkw == 1 and welkw_ladeleistung == 0:
                    welkw_ladeleistung = schnellladung
                    elkw_SOC_t = welkw_SOC + welkw_ladeleistung / 4 / welkw_S
                    if elkw_SOC_t > 1:
                        welkw_ladeleistung = (1 - welkw_SOC) * 4 * welkw_S
                    welkw_SOC = welkw_SOC + welkw_ladeleistung / 4 / welkw_S
                    ladeleistung_elkw1_a = welkw_ladeleistung
                    welkw_netzbezugauswaerts = welkw_netzbezugauswaerts + welkw_ladeleistung / 4
                    welkw_Energiebezug = welkw_Energiebezug + welkw_ladeleistung / 4
                ### unterwegs Akku tauschen bei raststätte/wechselstation
                if status_welkw == 1 and welkw_SOC < SOC_fuer_wechsel:
                    welkw_SOC = welkw_SOC + Kapazität_Wechselbatterie / welkw_S
                    welkw_Energiebezug = welkw_Energiebezug + Kapazität_Wechselbatterie
                    welkw_counter_wechsel = welkw_counter_wechsel + 1
                INFO_WELKW.iat[row, welkw_last] = last_S
                INFO_WELKW.iat[row, welkw_rueck] = ruecksp
                INFO_WELKW.iat[row, welkw_soc] = welkw_SOC
                if (weekday_S < 5 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") > datetime.datetime.strptime("07:00:00", "%H:%M:%S")
                    and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("20:00:00","%H:%M:%S")) \
                    or (weekday_S == 5 and datetime.datetime.strptime(str(zeit_S),"%H:%M:%S") > datetime.datetime.strptime("07:00:00", "%H:%M:%S")
                    and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("13:00:00", "%H:%M:%S")):
                    preis = welkw_ladeleistung_n / 4 * hochtarif
                else:
                    preis = welkw_ladeleistung_n / 4 * niedertarif
                welkw_summe_preis = welkw_summe_preis + preis
                LG_S.iat[row, LG_S.columns.get_loc("ANALYSE_SPITZE")] = last_S
        BILANZ_WELKW.iat[0, counter_title] = welkw_S
        BILANZ_WELKW.iat[1, counter_title] = Verbrauch
        BILANZ_WELKW.iat[2, counter_title] = welkw_Energiebezug
        BILANZ_WELKW.iat[3, counter_title] = round(welkw_Netzbezug)
        BILANZ_WELKW.iat[4, counter_title] = round(welkw_Eigenverbrauch)
        BILANZ_WELKW.iat[5, counter_title] = round(welkw_netzbezugauswaerts)
        BILANZ_WELKW.iat[6, counter_title] = elkw_summe_break_diff
        BILANZ_WELKW.iat[7, counter_title] = welkw_counter_wechsel
        BILANZ_WELKW.iat[8, counter_title] = round(welkw_counter_wechsel * Kapazität_Wechselbatterie)
        BILANZ_WELKW.iat[9, counter_title] = round(welkw_summe_preis)
        BILANZ_WELKW.iat[10, counter_title] = round(welkw_netzbezugauswaerts * tarif_schnellladen)
        BILANZ_WELKW.iat[11, counter_title] = round(welkw_Eigenverbrauch * rueckspeisungstarif)
        BILANZ_WELKW.iat[12, counter_title] = round(welkw_counter_wechsel * Kosten_Wechselbatterie)
        BILANZ_WELKW.iat[13, counter_title] = round(welkw_summe_preis + welkw_netzbezugauswaerts * tarif_schnellladen + welkw_Eigenverbrauch * rueckspeisungstarif +welkw_counter_wechsel * Kosten_Wechselbatterie)
        BILANZ_WELKW.iat[14, counter_title] = round(JahreskostenWELKW_var)
        BILANZ_WELKW.iat[15, counter_title] = round(Jahreskosten_dc_ladestation)
        BILANZ_WELKW.iat[16, counter_title] = round(JahreskostenWELKW_var + Jahreskosten_dc_ladestation+ round(welkw_summe_preis + welkw_netzbezugauswaerts * tarif_schnellladen + welkw_Eigenverbrauch * rueckspeisungstarif +welkw_counter_wechsel * Kosten_Wechselbatterie))
        counter_title = counter_title + 1
    for row in range(1, len(BILANZ_WELKW)):
        BILANZ_WELKW.iat[row, 4] = BILANZ_WELKW.iat[row, 1] + BILANZ_WELKW.iat[row, 2] + BILANZ_WELKW.iat[row, 3]

if EPKW_Sim ==1:
    names_Bilanz_EPKW = ["Energiebezug Speicher [kWh/a]", "Netzbezug [kWh/a]", "Eigenverbrauchserhöhung [kWh/a]",
                         "Kosten Netzbezug [CHF/a]", "Kosten Eigenverbrauch [CHF/a]", "Investition "+str(N_Ladestationen)+" Ladestationen [CHF]",
                         "Unterhalt "+str(N_Ladestationen)+" Ladestationen [CHF/a]", "Jahreskosten Inv + Unt [CHF/a]", "Jahreskosten total [CHF/a]"]
    BILANZ_EPKW = pd.DataFrame(index=range(len(names_Bilanz_EPKW)), columns=[output_excel_S])
    BILANZ_EPKW[output_excel_S] = names_Bilanz_EPKW
    BILANZ_EPKW["EPKW"] = None
    INFO_EPKW = pd.DataFrame(index=range(len(LG_S)), columns=columns)
    INFO_EPKW = INFO_EPKW.fillna(0)
    INFO_EPKW["Time"] = LG_S["Time"]
    INFO_EPKW["Weekday"] = LG_S["Weekday"]
    INFO_EPKW["Last_EPKW"] = LG_S["kW_PV_S"]
    INFO_EPKW["Ruecksp_EPKW"] = LG_S["kW_RS_S"]
    epkw_last = INFO_EPKW.columns.get_loc('Last_EPKW')
    epkw_rueck = INFO_EPKW.columns.get_loc('Ruecksp_EPKW')
    INFO_EPKW["EPKW_SOC"] = None
    sim_epkw_soc = INFO_EPKW.columns.get_loc("EPKW_SOC")
    Epkw_SOC = epkw_wochenstart_kWh
    epkw_Energiebezug = 0
    epkw_Netzbezug = 0
    epkw_Eigenverbrauch = 0
    for row in range(0, len(INFO_EPKW)):
        if ELKW_Sim ==0 and ELKW_Sim ==0:
            last_S = INFO_EPKW.iat[row, epkw_last]
            ruecksp = INFO_EPKW.iat[row, epkw_rueck]
        if ELKW_Sim >0:
            last_S = INFO_ELKW.iat[row, elkw_last]
            ruecksp = INFO_ELKW.iat[row, elkw_rueck]
        if WELKW_Sim >0:
            last_S = INFO_WELKW.iat[row, welkw_last]
            ruecksp = INFO_WELKW.iat[row, welkw_rueck]
        date_S = LG_S.iat[row, date]
        zeit_S = LG_S.iat[row, zeit]
        weekday_S = LG_S.iat[row, weekday]
        month_S = LG_S.iat[row, month]
        ladeleistung_pkw = 0
        ####### Epkw schlaufe
        if EPKW_Sim == 1:
            if weekday_S == 0 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") == datetime.datetime.strptime(
                    "00:00:00", "%H:%M:%S"):
                Epkw_SOC = epkw_wochenstart_kWh  ##SOC bei 50% für wochenstart
            if weekday_S <= 4 and datetime.datetime.strptime(str(zeit_S),"%H:%M:%S") >= datetime.datetime.strptime("08:00:00", "%H:%M:%S") \
                    and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("17:00:00","%H:%M:%S"):
                # laden aus überschuss pv
                if ruecksp > 0 and Epkw_SOC < 1:
                    if ruecksp > ladeleistung_stationen * N_Ladestationen:
                        ladeleistung_pkw = ladeleistung_stationen * N_Ladestationen
                        Epkw_SOC_t = Epkw_SOC + ladeleistung_pkw / N_Ladestationen / 4 / durchsch_kapazität
                        if Epkw_SOC_t > 1:
                            ladeleistung_pkw = (1 - Epkw_SOC) * 4 * N_Ladestationen * durchsch_kapazität
                    if ruecksp < ladeleistung_stationen * N_Ladestationen:
                        ladeleistung_pkw = ruecksp
                        Epkw_SOC_t = Epkw_SOC + ladeleistung_pkw / N_Ladestationen / 4 / durchsch_kapazität
                        if Epkw_SOC_t > 1:
                            ladeleistung_pkw = (1 - Epkw_SOC) * 4 * N_Ladestationen * durchsch_kapazität
                    ruecksp = ruecksp - ladeleistung_pkw
                    epkw_Energiebezug = epkw_Energiebezug + ladeleistung_pkw / 4
                    epkw_Eigenverbrauch = epkw_Eigenverbrauch + ladeleistung_pkw / 4
                    Epkw_SOC = Epkw_SOC + ladeleistung_pkw / N_Ladestationen / 4 / durchsch_kapazität
                if Epkw_SOC < 1 and ladeleistung_pkw < ladeleistung_stationen * N_Ladestationen and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") >= datetime.datetime.strptime("14:00:00", "%H:%M:%S") \
                        and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("17:00:00", "%H:%M:%S"):
                    #laden von netz
                    ladeleistung_pkw = ladeleistung_stationen/2 * N_Ladestationen - ladeleistung_pkw
                    Epkw_SOC_t = Epkw_SOC + ladeleistung_pkw / N_Ladestationen / 4 / durchsch_kapazität
                    if Epkw_SOC_t > 1:
                        ladeleistung_pkw = (1 - Epkw_SOC) * 4 * N_Ladestationen * durchsch_kapazität
                    last_S = last_S + ladeleistung_pkw
                    Epkw_SOC = Epkw_SOC + ladeleistung_pkw / N_Ladestationen / 4 / durchsch_kapazität
                    epkw_Energiebezug = epkw_Energiebezug + ladeleistung_pkw / 4
                    epkw_Netzbezug = epkw_Netzbezug + ladeleistung_pkw / 4
            if weekday_S > 0 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") == datetime.datetime.strptime("00:15:00","%H:%M:%S"):
                Epkw_SOC = (durchsch_kapazität - epkw_ladekapazitaet) / durchsch_kapazität
            INFO_EPKW.iat[row, epkw_last] = last_S
            INFO_EPKW.iat[row, epkw_rueck] = ruecksp
            INFO_EPKW.iat[row, sim_epkw_soc] = Epkw_SOC
            LG_S.iat[row, LG_S.columns.get_loc("ANALYSE_SPITZE")] = last_S
    BILANZ_EPKW.iat[0, 1] = round(epkw_Energiebezug)
    BILANZ_EPKW.iat[1, 1] = round(epkw_Netzbezug)
    BILANZ_EPKW.iat[2, 1] = round(epkw_Eigenverbrauch)
    BILANZ_EPKW.iat[3, 1] = round(epkw_Netzbezug * hochtarif)
    BILANZ_EPKW.iat[4, 1] = round(epkw_Eigenverbrauch * rueckspeisungstarif)
    BILANZ_EPKW.iat[5, 1] = round(ac_ladestation_inv)
    BILANZ_EPKW.iat[6, 1] = round(ac_ladestation_unterhalt)
    BILANZ_EPKW.iat[7, 1] = round(Jahreskosten_ac_ladestation)
    BILANZ_EPKW.iat[8, 1] = round(epkw_Netzbezug * hochtarif + epkw_Eigenverbrauch * rueckspeisungstarif + Jahreskosten_ac_ladestation)

if Sim_LG == 1:
    names_Bilanz_LG = ["Netzbezug [kWh/a]", "Thermoölbezug [kWh/a]", "mittlerer Netzbezug [kW/a]",
                       "mittlere monatliche Lastspitze [kW/a]",
                       "Kosten Netzbezug [CHF/a]", "Kosten Thermoöl [CHF/a]", "Kosten Lastspitzen [CHF/a]",
                       "Kosten total Lastgang bestehend, ohne weitere Komponenten [CHF/a]"]
    BILANZ_LG = pd.DataFrame(index=range(len(names_Bilanz_LG)), columns=[output_excel_S])
    BILANZ_LG[output_excel_S] = names_Bilanz_LG
    BILANZ_LG[LG_name] = None
    # max limit without PV
    ee = LG_S.groupby('month')['kW_Last'].transform(max) == LG_S['kW_Last']
    max_limit_source = LG_S[ee]
    counter = 1
    for row in range(0, 12):
        N = max_limit_source.iat[row, max_limit_source.columns.get_loc('month')] - counter
        if N < 0:
            index = max_limit_source.index
            max_limit_source = max_limit_source.drop(index[row])
        else:
            counter = counter + 1
        if counter == 13:
            break
    summe_preis_LG = 0
    for row in range(0, len(LG_S)):
        last_b_S = LG_S.iat[row, last_best]
        date_S = LG_S.iat[row, date]
        zeit_S = LG_S.iat[row, zeit]
        weekday_S = LG_S.iat[row, weekday]
        month_S = LG_S.iat[row, month]
        #### Bilanzierung von Einnahmen aus Eigenverbraucherhöhung durch die PV-Anlage
        if (weekday_S < 5 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") > datetime.datetime.strptime(
                "07:00:00", "%H:%M:%S") \
            and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("20:00:00",
                                                                                                  "%H:%M:%S")) \
                or (
                weekday_S == 5 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") > datetime.datetime.strptime(
            "07:00:00", "%H:%M:%S") \
                and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("13:00:00",
                                                                                                      "%H:%M:%S")):
            preis = last_b_S / 4 * hochtarif
        else:
            preis = last_b_S / 4 * niedertarif
        summe_preis_LG = summe_preis_LG + preis
    BILANZ_LG.iat[0, BILANZ_LG.columns.get_loc(LG_name)] = round(LG_S["kW_Last"].sum() / 4)
    if LG_name == "Effizient + Thermoöl":
        BILANZ_LG.iat[1, BILANZ_LG.columns.get_loc(LG_name)] = round(öl_verbrauch)
        BILANZ_LG.iat[5, BILANZ_LG.columns.get_loc(LG_name)] = round(öl_verbrauch * öl_tarif)
    else:
        BILANZ_LG.iat[1, BILANZ_LG.columns.get_loc(LG_name)] = 0
        BILANZ_LG.iat[5, BILANZ_LG.columns.get_loc(LG_name)] = 0
    BILANZ_LG.iat[2, BILANZ_LG.columns.get_loc(LG_name)] = round(LG_S["kW_Last"].mean())
    BILANZ_LG.iat[3, BILANZ_LG.columns.get_loc(LG_name)] = round(max_limit_source['kW_Last'].sum() / 12)
    BILANZ_LG.iat[4, BILANZ_LG.columns.get_loc(LG_name)] = round(summe_preis_LG)
    BILANZ_LG.iat[6, BILANZ_LG.columns.get_loc(LG_name)] = round(max_limit_source['kW_Last'].sum() * leistungsspitzenpreis)
    BILANZ_LG.iat[7, BILANZ_LG.columns.get_loc(LG_name)] = round(BILANZ_LG.iat[4, BILANZ_LG.columns.get_loc(LG_name)] + \
                                                           BILANZ_LG.iat[5, BILANZ_LG.columns.get_loc(LG_name)] + \
                                                           BILANZ_LG.iat[6, BILANZ_LG.columns.get_loc(LG_name)])

#PV bei S = 0
if len(Speicher)>0:
    if ELKW_Sim >0 or WELKW_Sim>0:
        for S in Speicher:
            del LG_S[S]
    from Analyse_V2 import LG_S, new_monthly_limit
    names_BILANZ_PV = ["Produktion [kWh/a]", "Volllaststunden [h]", "Eigenverbrauch [kWh/a]", "Rückspeisung [kWh/a]",
                       "Lastreduktion [kW/a]", "jährliche Einsparung_EV [CHF/a]",
                       "jährlicher Gewinn Rückspeisung [CHF/a]", "jährliche Einsparung_PS [CHF/a]",
                       "Einmalvergütung [CHF]", "Investitionskosten [CHF]", "jährliche Ausgaben [CHF/a]",
                       "jährliche Einnahmen [CHF/a]", "Amortisationszeit [a]", "Jahreskosten (25J) [CHF/a]",
                       "Business Case PV (25J) [CHF/a]"]
    BILANZ_PV = pd.DataFrame(index=range(len(names_BILANZ_PV)), columns=[output_excel_S])
    BILANZ_PV[output_excel_S] = names_BILANZ_PV
    BILANZ_PV[str(PV_neu)] = None
    names_BILANZ_BATTERIESPEICHER = ["Energiebezug Speicher [kWh/a]", "Volllastzyklen", "Netzbezug [kWh/a]",
                                     "jährliche Lastreduktion [kW/a]", "jährliche Eigenverbrauchserhöhung [kWh/a]",
                                     "jährliche Einsparung_PS [CHF/a]", "jährliche Einsparung_EV [CHF/a]",
                                     "Investitionskosten [CHF]", "Jahreskosten [CHF/a]", "Businesscase BS [CHF/a]"]
    BILANZ_BATTERIESPEICHER = pd.DataFrame(index=range(len(names_BILANZ_BATTERIESPEICHER)), columns=[output_excel_S])
    BILANZ_BATTERIESPEICHER[output_excel_S] = names_BILANZ_BATTERIESPEICHER
    PROFIL_SPEICHER = pd.DataFrame(index=range(len(LG_S)), columns=columns)
    PROFIL_SPEICHER = PROFIL_SPEICHER.fillna(0)
    PROFIL_SPEICHER["Time"] = LG_S["Time"]
    INFO_SPEICHER = pd.DataFrame(index=range(len(LG_S)), columns=columns)
    INFO_SPEICHER = INFO_SPEICHER.fillna(0)
    INFO_SPEICHER["Time"] = LG_S["Time"]
    INFO_SPEICHER["Weekday"] = LG_S["Weekday"]
    INFO_SPEICHER["kW_RS_S"] = LG_S["kW_RS_S"]
    SIM_MONTH = pd.DataFrame(index=range(3000),columns=["1_month", "2_month", "3_month", "4_month", "5_month", "6_month", "7_month","8_month", "9_month", "10_month", "11_month", "12_month"])
    SIM_MONTH = SIM_MONTH.fillna(0)
    MAX_MONTH = pd.DataFrame(index=range(12), columns=["month"])
    MAX_MONTH["month"] = month_titles
    LOAD_RED = pd.DataFrame(index=range(12), columns=["month"])
    LOAD_RED["month"] = month_titles
    LOAD_RED_PV = pd.DataFrame(index=range(12), columns=["month"])
    LOAD_RED_PV["month"] = month_titles
    LOAD_RED_PV[str(PV_neu)] = None
    sim_loadred_PV = LOAD_RED_PV.columns.get_loc(str(PV_neu))
    EINSPARUNG = pd.DataFrame(index=range(12), columns=["month"])
    EINSPARUNG["month"] = month_titles
    for S in Speicher:
        Sp_i = LG_S.columns.get_loc(S)
        ladestand = 0 #kWh
        SOC = 0
        Energiebezug = 0
        einsparungjahr = 0
        lastredjahr = 0
        summe_preis = 0
        summe_eigen = 0
        Netzbezug_tot = 0
        Ruecksp_tot = 0
        sum_max_load_month = 0
        summe_preis_netzbezug = 0
        summe_preis_netzbezug_pv = 0
        summe_netzbezug = 0
        if S == 0:
            summe_eigen_PV = 0
            summe_preis_PV = 0
        PROFIL_SPEICHER["Last_"+str(S)] = None
        sim_load = PROFIL_SPEICHER.columns.get_loc("Last_"+str(S))
        PROFIL_SPEICHER["Ruecksp_"+str(S)] = None
        sim_rueck = PROFIL_SPEICHER.columns.get_loc("Ruecksp_"+str(S))
        if S > 0:
            name_soc = str(S) + "_SOC"
            INFO_SPEICHER[name_soc] = None
            sim_soc = INFO_SPEICHER.columns.get_loc(name_soc)
            name_ladest = str(S) + "_Ladestand"
            INFO_SPEICHER[name_ladest] = None
            sim_ladest = INFO_SPEICHER.columns.get_loc(name_ladest)
            name_ladeleistung = str(S) + "_Ladeleistung_N"
            INFO_SPEICHER[name_ladeleistung] = None
            sim_ladeleistung = INFO_SPEICHER.columns.get_loc(name_ladeleistung)
            name_ladeleistung_PV = str(S) + "_Ladeleistung_PV"
            INFO_SPEICHER[name_ladeleistung_PV] = None
            sim_ladeleistung_PV = INFO_SPEICHER.columns.get_loc(name_ladeleistung_PV)
        name_max = str(S)
        MAX_MONTH[name_max] = None
        sim_max_loc = MAX_MONTH.columns.get_loc(name_max)
        name_loadred = str(S)
        LOAD_RED[name_loadred] = None
        sim_loadred = LOAD_RED.columns.get_loc(name_loadred)
        name_einsp = str(S)
        EINSPARUNG[name_einsp] = None
        sim_einsp = EINSPARUNG.columns.get_loc(name_einsp)
        name_res = "Batteriespeicher "+str(S)+"kWh"
        BILANZ_BATTERIESPEICHER[name_res] = None
        sim_res = BILANZ_BATTERIESPEICHER.columns.get_loc(name_res)
        name_tot = "Batteriespeicher "+str(S)+"kWh"
        BILANZ_TOTAL[name_tot] = None
        sim_tot = BILANZ_TOTAL.columns.get_loc(name_tot)
        counter1 = 0
        counter2 = 0
        counter3 = 0
        counter4 = 0
        counter5 = 0
        counter6 = 0
        counter7 = 0
        counter8 = 0
        counter9 = 0
        counter10 = 0
        counter11 = 0
        counter12 = 0
        for row in range(0, len(LG_S)):
            if ELKW_Sim ==0 and WELKW_Sim==0 and EPKW_Sim ==0:
                last_S = LG_S.iat[row, last]
                ruecksp = LG_S.iat[row, rueck]
            if ELKW_Sim > 0 and EPKW_Sim ==0:
                last_S = INFO_ELKW.iat[row, elkw_last]
                ruecksp = INFO_ELKW.iat[row, elkw_rueck]
            if WELKW_Sim > 0 and EPKW_Sim ==0:
                last_S = INFO_WELKW.iat[row, welkw_last]
                ruecksp = INFO_WELKW.iat[row, welkw_rueck]
            if EPKW_Sim ==1:
                last_S = INFO_EPKW.iat[row, epkw_last]
                ruecksp = INFO_EPKW.iat[row, epkw_rueck]
            last_pv = LG_S.iat[row, last]
            ruecksp_pv = LG_S.iat[row, rueck]
            speichergrenze_S = LG_S.iat[row, Sp_i]
            date_S = LG_S.iat[row, date]
            zeit_S = LG_S.iat[row, zeit]
            weekday_S = LG_S.iat[row, weekday]
            month_S = LG_S.iat[row, month]
            ladeleistung_N = 0
            ladeleistung_R = 0
    ####### Stationärer Speicher laden mit Netz oder von Rückspeisung PV
            if ladestand < S and last_S <= (Faktor_laden * speichergrenze_S) and S > 0:
                if datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") >= datetime.datetime.strptime("20:00:00", "%H:%M:%S") \
                        or datetime.datetime.strptime(str(zeit_S), "%H:%M:%S")  <= datetime.datetime.strptime("06:00:00", "%H:%M:%S"):
                    ladeleistung_N = Faktor_laden * speichergrenze_S - last_S
                    if ladeleistung_N > S:
                        ladeleistung_N = S
                if ruecksp > 0 and ruecksp > S:
                    ladeleistung_R = S
                    ladeleistung_N = 0
                    ruecksp = ruecksp - S
                if ruecksp > 0 and ruecksp < S:
                    ladeleistung_R = ruecksp
                    ladeleistung_N = 0
                    ruecksp = 0
                if weekday_S == 5 or weekday_S == 6 or weekday_S == 7:
                    ladeleistung_N = 0
                if weekday_S == 4 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") >= datetime.datetime.strptime("20:00:00", "%H:%M:%S"):
                    ladeleistung_N = 0
                if Eigenverbrauch > 2:
                    ladeleistung_N = 0
                ladestand = ladestand + ladeleistung_N/4 + ladeleistung_R/4
                if ladestand > S:
                    if ladeleistung_N > 0:
                        ladeleistung_N = ladeleistung_N - (ladestand-S)*4
                    if ladeleistung_R > 0:
                        ladeleistung_R = ladeleistung_R - (ladestand - S) * 4
                    ladestand = S
            last_S = last_S + ladeleistung_N
            PROFIL_SPEICHER.iat[row, sim_rueck] = ruecksp
            PROFIL_SPEICHER.iat[row, sim_load] = last_S
    ####### stationärer Speicher für Eigenverbrauhcserhöhung verwenden, je nach Einstellung in main
            if Eigenverbrauch > 0 and S>0:
                if last_S < Eigenverbrauch_grenze and Eigenverbrauch > 1:
                    if last_S > ladestand * 4:
                        last_S = last_S - ladestand * 4
                        Energiebezug = Energiebezug + ladestand
                        ladestand = 0
                    if last_S <= ladestand * 4:
                        Energiebezug = Energiebezug + last_S / 4
                        ladestand = ladestand - last_S / 4
                        last_S = 0
                    PROFIL_SPEICHER.iat[row, sim_load] = last_S
                if (weekday_S == 5 or weekday_S == 6 or weekday_S == 7) and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") >= datetime.datetime.strptime("07:00:00", "%H:%M:%S"):
                    if ladestand > 0:
                        if last_S > ladestand*4:
                            last_S = last_S - ladestand*4
                            Energiebezug = Energiebezug + ladestand
                            ladestand = 0
                        if last_S <= ladestand*4:
                            Energiebezug = Energiebezug + last_S/4
                            ladestand = ladestand - last_S/4
                            last_S = 0
                    PROFIL_SPEICHER.iat[row, sim_load] = last_S
            #Speicher entladen für Peak-Shaving
            if last_S > speichergrenze_S and S>0:
                if last_S - speichergrenze_S <= S:
                    if last_S - speichergrenze_S <= ladestand * 4:
                        ladestand = ladestand - (last_S - speichergrenze_S)/4
                        Energiebezug = Energiebezug + (last_S - speichergrenze_S)/4
                        last_S = speichergrenze_S
                    if last_S - speichergrenze_S >= ladestand * 4:
                        last_S = last_S - ladestand * 4
                        Energiebezug = Energiebezug + ladestand
                        ladestand = 0
                if last_S - speichergrenze_S > S:
                    if S <= ladestand * 4:
                        ladestand = ladestand - S / 4
                        Energiebezug = Energiebezug + S / 4
                        last_S = last_S - S
                    if S >= ladestand * 4:
                        last_S = last_S - ladestand * 4
                        Energiebezug = Energiebezug + ladestand
                        ladestand = 0
                PROFIL_SPEICHER.iat[row, sim_load] = last_S
            # Bilanzierung von Einnahmen aus Eigenverbraucherhöhung durch den Speicher
            if S > 0:
                INFO_SPEICHER.iat[row, sim_ladeleistung] = ladeleistung_N
                INFO_SPEICHER.iat[row, sim_ladeleistung_PV] = ladeleistung_R
                SOC = 1 / S * ladestand
                INFO_SPEICHER.iat[row, sim_soc] = SOC
                INFO_SPEICHER.iat[row, sim_ladest] = ladestand
                eigenverbr_erh = ladeleistung_R/4
                netzbezug_15min = ladeleistung_N/4
                if (weekday_S < 5 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") > datetime.datetime.strptime("07:00:00", "%H:%M:%S")
                    and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("20:00:00", "%H:%M:%S")) \
                        or (weekday_S == 5 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") > datetime.datetime.strptime("07:00:00", "%H:%M:%S")
                            and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("13:00:00", "%H:%M:%S")):
                    preis = ladeleistung_N/4 * hochtarif
                    preis_netzbezug = last_S/4 *hochtarif
                else:
                    preis = ladeleistung_N/4 * niedertarif
                    preis_netzbezug = last_S/4*niedertarif
                summe_preis = summe_preis + preis
                summe_eigen = summe_eigen + eigenverbr_erh
                summe_netzbezug = summe_netzbezug + netzbezug_15min
                summe_preis_netzbezug = summe_preis_netzbezug + preis_netzbezug
            # Bilanzierung von Einnahmen aus Eigenverbraucherhöhung durch die PV-Anlage
            if S == 0:
                eigenverbr_erh = (round(LG_S.iat[row, LG_S.columns.get_loc("kW_Last")]) - LG_S.iat[row, last])/4 # von LG_S der Netzbezug
                summe_eigen_PV = summe_eigen_PV + eigenverbr_erh
                if (weekday_S < 5 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") > datetime.datetime.strptime("07:00:00", "%H:%M:%S")
                    and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("20:00:00", "%H:%M:%S")) \
                        or (weekday_S == 5 and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") > datetime.datetime.strptime("07:00:00", "%H:%M:%S")
                            and datetime.datetime.strptime(str(zeit_S), "%H:%M:%S") <= datetime.datetime.strptime("13:00:00", "%H:%M:%S")):
                    preis = eigenverbr_erh * hochtarif
                    preis_netzbezug = last_S / 4 * hochtarif
                    preis_netzbezug_pv = last_pv / 4 * hochtarif
                else:
                    preis = eigenverbr_erh * niedertarif
                    preis_netzbezug = last_S / 4 * niedertarif
                    preis_netzbezug_pv = last_pv / 4 * hochtarif
                summe_preis_PV = summe_preis_PV + preis
                summe_preis_netzbezug = summe_preis_netzbezug + preis_netzbezug
                summe_preis_netzbezug_pv = summe_preis_netzbezug_pv + preis_netzbezug_pv
        # Monate aufteilen / diese auflistung wird nicht ausgegeben // für maximum pro monat
            last_SM = last_S
            if S ==0:
                last_SM = LG_S.iat[row, last]
            if month_S == 1:
                name_m = str(month_S) + "_month"
                sim_month1 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter1, sim_month1] = last_SM
                counter1 = counter1 + 1
            if month_S == 2:
                name_m = str(month_S) + "_month"
                sim_month2 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter2, sim_month2] = last_SM
                counter2 = counter2 + 1
            if month_S == 3:
                name_m = str(month_S) + "_month"
                sim_month3 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter3, sim_month3] = last_SM
                counter3 = counter3 + 1
            if month_S == 4:
                name_m = str(month_S) + "_month"
                sim_month4 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter4, sim_month4] = last_SM
                counter4 = counter4 + 1
            if month_S == 5:
                name_m = str(month_S) + "_month"
                sim_month5 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter5, sim_month5] = last_SM
                counter5 = counter5 + 1
            if month_S == 6:
                name_m = str(month_S) + "_month"
                sim_month6 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter6, sim_month6] = last_SM
                counter6 = counter6 + 1
            if month_S == 7:
                name_m = str(month_S) + "_month"
                sim_month7 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter7, sim_month7] = last_SM
                counter7 = counter7 + 1
            if month_S == 8:
                name_m = str(month_S) + "_month"
                sim_month8 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter8, sim_month8] = last_SM
                counter8 = counter8 + 1
            if month_S == 9:
                name_m = str(month_S) + "_month"
                sim_month9 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter9, sim_month9] = last_SM
                counter9 = counter9 + 1
            if month_S == 10:
                name_m = str(month_S) + "_month"
                sim_month10 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter10, sim_month10] = last_SM
                counter10 = counter10 + 1
            if month_S == 11:
                name_m = str(month_S) + "_month"
                sim_month11 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter11, sim_month11] = last_SM
                counter11 = counter11 + 1
            if month_S == 12:
                name_m = str(month_S) + "_month"
                sim_month12 = SIM_MONTH.columns.get_loc(name_m)
                SIM_MONTH.iat[counter12, sim_month12] = last_SM
                counter12 = counter12 + 1
            Netzbezug_tot = Netzbezug_tot+last_S/4
            Ruecksp_tot = Ruecksp_tot + ruecksp/4
        #  Maximum Spitze pro monat ausgeben
        for column in range(0, len(SIM_MONTH)):
            #max_load_month = SIM_MONTH.max(0, column)
            max_load_month = SIM_MONTH.max()
            counterM = 0
            for row in range(0, len(MAX_MONTH)):
                MAX_MONTH.iat[row, sim_max_loc] = max_load_month[counterM]
                if S == 0:
                    LOAD_RED_PV.iat[row, sim_loadred_PV] = max_limit_source.iat[row, max_limit_source.columns.get_loc('kW_Last')] - max_load_month[counterM]
                    LOAD_RED.iat[row, sim_loadred] = 0
                counterM = counterM + 1
                if S > 0:
                    calc = MAX_MONTH.iat[row, MAX_MONTH.columns.get_loc("0")] - MAX_MONTH.iat[row, sim_max_loc]
                    if calc <= LOAD_RED.iat[row, counterLR]:
                        calc = LOAD_RED.iat[row, counterLR]
                    LOAD_RED.iat[row, sim_loadred] = calc
                    EINSPARUNG.iat[row, sim_einsp] = calc * leistungsspitzenpreis
            sum_max_load_month = MAX_MONTH[name_max].sum()
        counterLR = counterLR + 1
        Jahreskosten_BS = 0
        #  Businesscase und Resulatet berechnen
        if S == 0: ##PV Reiter: ["Produktion [kWh/a]", "Volllaststunden [h]", "Eigenverbrauch [kWh/a]", "Rückspeisung [kWh/a]", "Lastreduktion [kW/a]", "jährliche Einsparung_EV [CHF/a]", "jährlicher Gewinn Rückspeisung [CHF/a]", "jährliche Einsparung_PS [CHF/a]", "Investitionskosten [CHF]", "jährliche Ausgaben [CHF/a]", "jährliche Einnahmen [CHF/a]", "Amortisationszeit [a]", "Jahreskosten (25J) [CHF/a]", "Business Case PV (25J) [CHF/a]"]
            einsparung_pv_ev = summe_preis_PV
            erhöhung_ruecks = LG_S["kW_RS_S"].sum()/4
            Produktion = (LG_S['kW_Last'].sum() - LG_S["kW_PV_S"].sum())/4 + erhöhung_ruecks
            BILANZ_PV.iat[0, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(Produktion)
            BILANZ_PV.iat[1, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(Produktion / PV_neu)
            BILANZ_PV.iat[2, BILANZ_PV.columns.get_loc(str(PV_neu))] = round((LG_S['kW_Last'].sum() - LG_S["kW_PV_S"].sum()) / 4)
            BILANZ_PV.iat[3, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(erhöhung_ruecks)
            BILANZ_PV.iat[4, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(LOAD_RED_PV[str(PV_neu)].sum())
            BILANZ_PV.iat[5, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(einsparung_pv_ev)
            BILANZ_PV.iat[6, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(erhöhung_ruecks * rueckspeisungstarif)
            BILANZ_PV.iat[7, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(LOAD_RED_PV[str(PV_neu)].sum() * leistungsspitzenpreis)
            BILANZ_PV.iat[8, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(EIV)
            BILANZ_PV.iat[9, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(Capex_PV - EIV)
            BILANZ_PV.iat[10, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(jährliche_Ausgaben_PV)
            jährliche_Einnahmen_PV = einsparung_pv_ev + erhöhung_ruecks * rueckspeisungstarif + LOAD_RED_PV[str(PV_neu)].sum() * leistungsspitzenpreis
            BILANZ_PV.iat[11, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(jährliche_Einnahmen_PV)
            BILANZ_PV.iat[12, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(Capex_PV / (jährliche_Einnahmen_PV - jährliche_Ausgaben_PV),1)
            BILANZ_PV.iat[13, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(JahreskostenPV_var)
            BC_PV = einsparung_pv_ev + erhöhung_ruecks * rueckspeisungstarif + LOAD_RED_PV[str(PV_neu)].sum() * leistungsspitzenpreis - JahreskostenPV_var
            BILANZ_PV.iat[14, BILANZ_PV.columns.get_loc(str(PV_neu))] = round(BC_PV)
            #BILANZ_TOTAL.iat[0, sim_tot] = round(LOAD_RED_PV[str(PV_neu)].sum())
            #BILANZ_TOTAL.iat[1, sim_tot] = (LG_S["kW_Last"].sum() - LG_S["kW_PV_S"].sum()) / 4
            #BILANZ_TOTAL.iat[2, sim_tot] = LOAD_RED_PV[str(PV_neu)].sum() * leistungsspitzenpreis
            #BILANZ_TOTAL.iat[3, sim_tot] = einsparung_pv_ev
            #BILANZ_TOTAL.iat[4, sim_tot] = LG_S["kW_RS_S"].sum() / 4 * rueckspeisungstarif
            #BILANZ_TOTAL.iat[5, sim_tot] = JahreskostenPV_var
            #BILANZ_TOTAL.iat[6, sim_tot] = BC_PV
        if S > 0:
            #Jahreskosten_BS = (Capex_BS_var * S + Capex_BS_fix)*A_BS+BU_BS
            Jahreskosten_BS = (-0.3617*S**2 + 1009.9*S + 6812.5) * (1 - Faktor_vergünstigung) * (A_BS+0.01)
            #Jahreskosten_BS = (Capex_BS_var * S + Capex_BS_fix)/N_BS+BU_BS
            lastredjahr = LOAD_RED[name_loadred].sum()
            einsparungjahr = lastredjahr * leistungsspitzenpreis
            BILANZ_BATTERIESPEICHER.iat[0, sim_res] = Energiebezug
            BILANZ_BATTERIESPEICHER.iat[1, sim_res] = Energiebezug / S
            BILANZ_BATTERIESPEICHER.iat[2, sim_res] = summe_netzbezug
            BILANZ_BATTERIESPEICHER.iat[3, sim_res] = lastredjahr
            BILANZ_BATTERIESPEICHER.iat[4, sim_res] = summe_eigen
            BILANZ_BATTERIESPEICHER.iat[5, sim_res] = einsparungjahr
            einsparungjahr_EV = summe_preis - INFO_SPEICHER[name_ladeleistung_PV].sum()/4*rueckspeisungstarif
            BILANZ_BATTERIESPEICHER.iat[6, sim_res] = einsparungjahr_EV
            BILANZ_BATTERIESPEICHER.iat[7, sim_res] = (-0.3617 * S ** 2 + 1009.9 * S + 6812.5) * (1 - Faktor_vergünstigung)
            BILANZ_BATTERIESPEICHER.iat[8, sim_res] = Jahreskosten_BS
            BILANZ_BATTERIESPEICHER.iat[9, sim_res] = einsparungjahr + einsparungjahr_EV - Jahreskosten_BS
            #BILANZ_TOTAL.iat[0, sim_tot] = LOAD_RED_PV[str(PV_neu)].sum() + lastredjahr
            #BILANZ_TOTAL.iat[1, sim_tot] = (LG_S["kW_Last"].sum() - LG_S["kW_PV_S"].sum()) / 4 + summe_eigen
            #BILANZ_TOTAL.iat[2, sim_tot] = LOAD_RED_PV[str(PV_neu)].sum() * leistungsspitzenpreis + einsparungjahr
            #BILANZ_TOTAL.iat[3, sim_tot] = einsparung_pv_ev + einsparungjahr_EV
            #BILANZ_TOTAL.iat[4, sim_tot] = LG_S["kW_RS_S"].sum() / 4 * rueckspeisungstarif - INFO_SPEICHER[name_ladeleistung_PV].sum() / 4 * rueckspeisungstarif
            #BILANZ_TOTAL.iat[5, sim_tot] = JahreskostenPV_var + Jahreskosten_BS
            #BILANZ_TOTAL.iat[6, sim_tot] = einsparungjahr + einsparungjahr_EV - Jahreskosten_BS + BC_PV
        x = 0
        BILANZ_TOTAL.iat[0, sim_tot] = round(Netzbezug_tot)
        BILANZ_TOTAL.iat[1+x, sim_tot] = round(Ruecksp_tot)
        BILANZ_TOTAL.iat[2, sim_tot] = round(sum_max_load_month)
        if ELKW_Sim > 1:
            BILANZ_TOTAL.iat[3, sim_tot] = round(BILANZ_ELKW.iat[5,1]+BILANZ_ELKW.iat[5,2]+BILANZ_ELKW.iat[5,3])
        elif WELKW_Sim > 1:
            BILANZ_TOTAL.iat[3, sim_tot] = round(BILANZ_WELKW.iat[5, 1] + BILANZ_WELKW.iat[5, 2] + BILANZ_WELKW.iat[5, 3] + BILANZ_WELKW.iat[8, 1] + BILANZ_WELKW.iat[8, 2] + BILANZ_WELKW.iat[8, 3])
        else:
            BILANZ_TOTAL.iat[3, sim_tot] = 0
        BILANZ_TOTAL.iat[4, sim_tot] = round((LG_S['kW_Last'].sum() - LG_S["kW_PV_S"].sum()) / 4 + erhöhung_ruecks-Ruecksp_tot)
        BILANZ_TOTAL.iat[5, sim_tot] = round(max_limit_source['kW_Last'].sum() - sum_max_load_month)
        BILANZ_TOTAL.iat[6, sim_tot] = round(summe_preis_netzbezug)

        BILANZ_TOTAL.iat[7, sim_tot] = round(sum_max_load_month * leistungsspitzenpreis)
        if ELKW_Sim > 1:
            BILANZ_TOTAL.iat[8, sim_tot] = round(BILANZ_ELKW.iat[9,1]+BILANZ_ELKW.iat[9,2]+BILANZ_ELKW.iat[9,3]) #kosten auswärtsladen elkw
        elif WELKW_Sim >1:
            BILANZ_TOTAL.iat[8, sim_tot] = round(BILANZ_WELKW.iat[10, 1] + BILANZ_WELKW.iat[10, 2] + BILANZ_WELKW.iat[10, 3] + BILANZ_WELKW.iat[12, 1] + BILANZ_WELKW.iat[12, 2] + BILANZ_WELKW.iat[12, 3])
        else:
            BILANZ_TOTAL.iat[8, sim_tot] = 0
        BILANZ_TOTAL.iat[9, sim_tot] = BILANZ_TOTAL.iat[6, sim_tot] + BILANZ_TOTAL.iat[7, sim_tot] + BILANZ_TOTAL.iat[8, sim_tot] #Stromkosten
        BILANZ_TOTAL.iat[10, sim_tot] = round(-Ruecksp_tot * rueckspeisungstarif)
        BILANZ_TOTAL.iat[11, sim_tot] = BILANZ_LG.iat[5,1]
        if ELKW_Sim >0:
            BILANZ_TOTAL.iat[12, sim_tot] = round(JahreskostenPV_var + Jahreskosten_BS + BILANZ_ELKW.iat[12,4] + BILANZ_ELKW.iat[13,4] + BILANZ_EPKW.iat[7,1])
        elif WELKW_Sim >1:
            BILANZ_TOTAL.iat[12, sim_tot] = round(JahreskostenPV_var + Jahreskosten_BS + BILANZ_WELKW.iat[14, 4] + BILANZ_WELKW.iat[15, 4] + BILANZ_EPKW.iat[7,1])
        else:
            BILANZ_TOTAL.iat[12, sim_tot] = round(JahreskostenPV_var + Jahreskosten_BS + JahreskostenLKW_var + BILANZ_EPKW.iat[7,1])
        BILANZ_TOTAL.iat[13, sim_tot] = BILANZ_TOTAL.iat[9, sim_tot] + BILANZ_TOTAL.iat[10, sim_tot] + BILANZ_TOTAL.iat[11, sim_tot] + BILANZ_TOTAL.iat[12, sim_tot]  # Kostentotal variante


print('Simulation Done!')