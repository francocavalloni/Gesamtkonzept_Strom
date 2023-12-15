import pandas as pd
from __main__ import Input_File
from read_LG_PV_V1 import date_titles, title, data_S
from calc_BC_V1 import öl_verbrauch, öl_tarif, JahreskostenLKW_var,  \
    lkw_inv, lkw_betriebskosten, lkw_Energiekosten,lkw_steuern, lkw_strecke, N_lkw
if Input_File:
    from read_input_excel import *
else:
    from alternative_input import *

"In diesem Skript werden die Dataframes für die Resultate vorbereitet."


# Vorbereiten des Lastgangs für die Simulation
LG_S = data_S
LG_S["ANALYSE_SPITZE"] = LG_S["kW_PV_A"]
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
# max limit without PV
max_limit_source = LG_S[LG_S.groupby('month')['kW_Last'].transform(max) == LG_S['kW_Last']]

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
names_BILANZ_TOTAL = [totname1, totname2, totname3, totname4, totname5, totname6, totname7, totname8, totname9, totname10, totname11, totname12, totname13, totname14]
BILANZ_TOTAL = pd.DataFrame(index=range(len(names_BILANZ_TOTAL)), columns=[output_excel_S])
BILANZ_TOTAL[output_excel_S] = names_BILANZ_TOTAL

#Set Dataframes ELKW
names_Bilanz_ELKW = ["Speicherkapazität [kWh]", "Verbrauch [kWh/km]", "Energiebezug Speicher [kWh/a]", "Netzbezug [kWh/a]", "Eigenverbrauchserhöhung [kWh/a]",
                     "Auswärtsladen [kWh/a]",  "Mittagspause Verlängert um [min]", "Häufigkeit Pause an Raststätte 30min laden",
                    "Kosten Netzbezug [CHF/a]","Kosten Laden Auswärts [CHF/a]","Kosten Eigenverbrauch [CHF/a]",
                     "Summe Energiekosten [CHF/a]", "Jahreskosten Inv + Unt [CHF/a]", "Jahreskosten Ladestation DC Inv + Unt", "Kosten total ELKW [CHF/a]"]
BILANZ_ELKW = pd.DataFrame(index=range(len(names_Bilanz_ELKW)), columns=[output_excel_S])
BILANZ_ELKW[output_excel_S] = names_Bilanz_ELKW
BILANZ_ELKW["ELKW1"] = None
BILANZ_ELKW["ELKW2"] = None
BILANZ_ELKW["ELKW3"] = None
BILANZ_ELKW["Summe 3 ELKWs"] = None
BILANZ_ELKW = BILANZ_ELKW.fillna(0)
INFO_ELKW = pd.DataFrame(index=range(len(LG_S)), columns=columns)
INFO_ELKW["Time"] = LG_S["Time"]
INFO_ELKW["Weekday"] = LG_S["Weekday"]
INFO_ELKW["ELKW1_Warnung"] = None
INFO_ELKW["ELKW1_SOC"] = None
INFO_ELKW["ELKW2_Warnung"] = None
INFO_ELKW["ELKW2_SOC"] = None
INFO_ELKW["ELKW3_Warnung"] = None
INFO_ELKW["ELKW3_SOC"] = None

#Set Dataframes WELKW
names_Bilanz_WELKW = ["Speicherkapazität [kWh]", "Verbrauch [kWh/km]", "Energiebezug Speicher [kWh/a]", "Netzbezug [kWh/a]", "Eigenverbrauchserhöhung [kWh/a]",
                     "Auswärtsladen [kWh/a]",  "Mittagspause Verlängert um [min]", "Häufigkeit Batteriewechsel an Station [5min]", "Laden durch Batteriewechsel [kWh/a]",
                    "Kosten Netzbezug [CHF/a]","Kosten Laden Auswärts [CHF/a]","Kosten Eigenverbrauch [CHF/a]", "Kosten Batteriewechsel [CHF/a",
                     "Summe Energiekosten [CHF/a]", "Jahreskosten Inv + Unt [CHF/a]", "Jahreskosten Ladestation DC Inv + Unt" ,  "Kosten total ELKW [CHF/a]"]
BILANZ_WELKW = pd.DataFrame(index=range(len(names_Bilanz_WELKW)), columns=[output_excel_S])
BILANZ_WELKW[output_excel_S] = names_Bilanz_WELKW
BILANZ_WELKW["WELKW1"] = None
BILANZ_WELKW["WELKW2"] = None
BILANZ_WELKW["WELKW3"] = None
BILANZ_WELKW["Summe 3 WELKWs"] = None
BILANZ_WELKW = BILANZ_WELKW.fillna(0)
INFO_WELKW = pd.DataFrame(index=range(len(LG_S)), columns=columns)
INFO_WELKW["Time"] = LG_S["Time"]
INFO_WELKW["Weekday"] = LG_S["Weekday"]
INFO_WELKW["WELKW1_Warnung"] = None
INFO_WELKW["WELKW1_SOC"] = None
INFO_WELKW["WELKW2_Warnung"] = None
INFO_WELKW["WELKW2_SOC"] = None
INFO_WELKW["WELKW3_Warnung"] = None
INFO_WELKW["WELKW3_SOC"] = None

#Set Dataframe EPKW
names_Bilanz_EPKW = ["Energiebezug Speicher [kWh/a]", "Netzbezug [kWh/a]", "Eigenverbrauchserhöhung [kWh/a]",
                     "Kosten Netzbezug [CHF/a]", "Kosten Eigenverbrauch [CHF/a]", "Investition "+str(N_Ladestationen)+" Ladestationen [CHF]",
                     "Unterhalt "+str(N_Ladestationen)+" Ladestationen [CHF/a]", "Jahreskosten Inv + Unt [CHF/a]", "Jahreskosten total [CHF/a]"]
BILANZ_EPKW = pd.DataFrame(index=range(len(names_Bilanz_EPKW)), columns=[output_excel_S])
BILANZ_EPKW[output_excel_S] = names_Bilanz_EPKW
BILANZ_EPKW["EPKW"] = None
BILANZ_EPKW = BILANZ_EPKW.fillna(0)
INFO_EPKW = pd.DataFrame(index=range(len(LG_S)), columns=columns)
INFO_EPKW = INFO_EPKW.fillna(0)
INFO_EPKW["Time"] = LG_S["Time"]
INFO_EPKW["Weekday"] = LG_S["Weekday"]
INFO_EPKW["Last_EPKW"] = LG_S["kW_PV_S"]
INFO_EPKW["Ruecksp_EPKW"] = LG_S["kW_RS_S"]
INFO_EPKW["EPKW_SOC"] = None

#Set Dataframe LG
names_Bilanz_LG = ["Netzbezug [kWh/a]", "Thermoölbezug [kWh/a]", "mittlerer Netzbezug [kW/a]",
                   "mittlere monatliche Lastspitze [kW/a]",
                   "Kosten Netzbezug [CHF/a]", "Kosten Thermoöl [CHF/a]", "Kosten Lastspitzen [CHF/a]",
                   "Kosten total Lastgang bestehend, ohne weitere Komponenten [CHF/a]"]
BILANZ_LG = pd.DataFrame(index=range(len(names_Bilanz_LG)), columns=[output_excel_S])
BILANZ_LG[output_excel_S] = names_Bilanz_LG
LG_name = "Bilanz vor Simulation"
BILANZ_LG[LG_name] = None
BILANZ_LG = BILANZ_LG.fillna(0)

#Set Dataframe PV
names_BILANZ_PV = ["Produktion [kWh/a]", "Volllaststunden [h]", "Eigenverbrauch [kWh/a]", "Rückspeisung [kWh/a]",
                   "Lastreduktion [kW/a]", "jährliche Einsparung_EV [CHF/a]",
                   "jährlicher Gewinn Rückspeisung [CHF/a]", "jährliche Einsparung_PS [CHF/a]",
                   "Einmalvergütung [CHF]", "Investitionskosten [CHF]", "jährliche Ausgaben [CHF/a]",
                   "jährliche Einnahmen [CHF/a]", "Amortisationszeit [a]", "Jahreskosten (25J) [CHF/a]",
                   "Business Case PV (25J) [CHF/a]"]
BILANZ_PV = pd.DataFrame(index=range(len(names_BILANZ_PV)), columns=[output_excel_S])
BILANZ_PV[output_excel_S] = names_BILANZ_PV
BILANZ_PV[str(PV_neu)] = None
BILANZ_PV = BILANZ_PV.fillna(0)

#Set Dataframe BS
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
SIM_MONTH = pd.DataFrame(index=range(3000),
                         columns=["1_month", "2_month", "3_month", "4_month", "5_month", "6_month", "7_month",
                                  "8_month", "9_month", "10_month", "11_month", "12_month"])
SIM_MONTH = SIM_MONTH.fillna(0)
MAX_MONTH = pd.DataFrame(index=range(12), columns=["month"])
MAX_MONTH["month"] = month_titles
LOAD_RED = pd.DataFrame(index=range(12), columns=["month"])
LOAD_RED["month"] = month_titles
LOAD_RED_PV = pd.DataFrame(index=range(12), columns=["month"])
LOAD_RED_PV["month"] = month_titles
LOAD_RED_PV[str(PV_neu)] = None
EINSPARUNG = pd.DataFrame(index=range(12), columns=["month"])
EINSPARUNG["month"] = month_titles

#Set Dataframe Eingabe Parameter und LKW Vergleich
Parameters_data = {
    'LG_name': output_excel_S,
    'Leistungsspitzenpreis [CHF/kW]': leistungsspitzenpreis,
    'Tarifvariation [%]': "+"+str(round((Tar_var-1)*100)),
    'Hochtarif [CHF/kWh]': hochtarif,
    'Niedertarif [CHF/kWh]': niedertarif,
    'Rückspeisungstarif [CHF/kWh]': rueckspeisungstarif,
    'PV kWp [kW]': PV_neu}
if thermooel_Sim:
    Parameters_data["Öl_verbrauch [kWh/a]"] = round(öl_verbrauch)
    Parameters_data["Öl_tarif [CHF/kWh]"] = öl_tarif
if SP_Sim:
    Parameters_data["Speicherkapazitäten [kWh]"] = Speicher
    Parameters_data["Sicherheitsfaktor Speicher"] = Faktor_Grenze
reshaped_para = {
    'Bezeichnung': [key for key in Parameters_data.keys()],
    'Wert': [value for value in Parameters_data.values()]}
PARAMETERS = pd.DataFrame(reshaped_para)

Diesel_data = {
    'Investitionskosten 3 LKWs [CHF]': round(lkw_inv),
    'Unterhaltskosten [CHF/a]': round(lkw_betriebskosten),
    "Energiekosten [CHF/a]" : round(lkw_Energiekosten),
    "Steuern [CHF/a]" : round(lkw_steuern),
    'Total Strecke [km/a]': round(lkw_strecke),
    "Lebensdauer [a]" : round(N_lkw),
    'Jahreskosten [CHF/a]': round(JahreskostenLKW_var),}
reshaped_diesel = {
    'Bezeichnung': [key for key in Diesel_data.keys()],
    'Wert': [value for value in Diesel_data.values()]}
BILANZ_DIESEL = pd.DataFrame(reshaped_diesel)