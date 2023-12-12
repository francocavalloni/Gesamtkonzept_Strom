import pandas as pd
import sys
import os
sys.path.insert(1, 'Module')
current_directory = os.getcwd()

"Inputs aus excel einlesen"
df_input = pd.read_excel('input_parameter.xlsx',1)
variable_names = df_input['Variabel'].tolist()
variable_values = df_input['Wert'].tolist()
input_dict = dict(zip(variable_names, variable_values))
#variabeln
LG_Sim = eval(input_dict.get('LG_Sim', 0))
PV_Sim = eval(input_dict.get('PV_Sim', 0))
ELKW_Sim = input_dict.get('ELKW_Sim', 0)
WELKW_Sim = input_dict.get('WELKW_Sim', 0)
EPKW_Sim = eval(input_dict.get('EPKW_Sim', 0))
SP_Sim = eval(input_dict.get('SP_Sim', 0))
excel = eval(input_dict.get('excel', 0))
plot = eval(input_dict.get('plot', 0))
LKW_Sim = False

"Paths aus excel einlesen"
df_paths = pd.read_excel('input_parameter.xlsx',2)
variable_names = df_paths['Variabel'].tolist()
variable_values = df_paths['Eingabe Path'].tolist()
variable_page = df_paths['Eingabe Page'].tolist()
paths_dict = dict(zip(variable_names, variable_values))
paths_page_dict = dict(zip(variable_names, variable_page))
output_excel_S = paths_dict.get('output_excel_S', 0)  #Ausgabedatei
r_DIR = paths_dict.get('r_DIR', 0)
r_LG_S = pd.read_excel(os.path.join(current_directory, r_DIR, paths_dict.get('r_LG_S', 0)), paths_page_dict.get('r_LG_S', 0)-1) #Lastgang 1
r_LG_A = pd.read_excel(os.path.join(current_directory, r_DIR, paths_dict.get('r_LG_A', 0)), paths_page_dict.get('r_LG_A', 0)-1) #Lastgang 1
r_PV_F = pd.read_excel(os.path.join(current_directory,r_DIR,paths_dict.get('r_PV', 0)), 1) #PV_Fassaden
r_PV_DS = pd.read_excel(os.path.join(current_directory,r_DIR,paths_dict.get('r_PV', 0)), 2) #PV_Dach_Süd
r_PV_DOW = pd.read_excel(os.path.join(current_directory,r_DIR,paths_dict.get('r_PV', 0)), 3) #PV_Dach_OstWest
r_PV_CP = pd.read_excel(os.path.join(current_directory,r_DIR,paths_dict.get('r_PV', 0)), 4) #PV_Dach_carport
r_LKW1 = pd.read_excel(os.path.join(current_directory,r_DIR,paths_dict.get('r_LKW', 0)), 0)
r_LKW2 = pd.read_excel(os.path.join(current_directory,r_DIR,paths_dict.get('r_LKW', 0)), 1)
r_LKW3 = pd.read_excel(os.path.join(current_directory,r_DIR,paths_dict.get('r_LKW', 0)), 2)
Feiertage = ["01-01","01-02", "12-25", "12-26"]

"Tarife aus excel einlesen"
df_tarife = pd.read_excel('input_parameter.xlsx',3)
variable_names = df_tarife['Variabel'].tolist()
variable_values = df_tarife['Wert'].tolist()
tarife_dict = dict(zip(variable_names, variable_values))
"##### TARIFE #####"
Tar_var = tarife_dict.get('Tar_var', 0)
mwst = tarife_dict.get('mwst', 0)
leistungsspitzenpreis = tarife_dict.get('leistungsspitzenpreis', 0)
hochtarif = tarife_dict.get('hochtarif', 0) * (mwst+1)
niedertarif = tarife_dict.get('niedertarif', 0) * (mwst+1)
rueckspeisungstarif = tarife_dict.get('rueckspeisungstarif', 0)
tarif_schnellladen = tarife_dict.get('tarif_schnellladen', 0)
Zins_WACC = tarife_dict.get('Zins_WACC', 0)
zeit_hochtarif_woche_start = tarife_dict.get('zeit_hochtarif_woche_start', 0)
zeit_hochtarif_woche_ende = tarife_dict.get('zeit_hochtarif_woche_ende', 0)
zeit_hochtarif_samstag_start = tarife_dict.get('zeit_hochtarif_samstag_start', 0)
zeit_hochtarif_samstag_ende = tarife_dict.get('zeit_hochtarif_samstag_ende', 0)

"Photovoltaik aus excel einlesen"
df_pv = pd.read_excel('input_parameter.xlsx',4)
variable_names = df_pv['Variabel'].tolist()
variable_values = df_pv['Wert'].tolist()
pv_dict = dict(zip(variable_names, variable_values))
dach_süd = pv_dict.get('dach_süd', 0)
dach_ostwest = pv_dict.get('dach_ostwest', 0)
fassade_süd = pv_dict.get('fassade_süd', 0)
fassade_west = pv_dict.get('fassade_west', 0)
fassade_nord = pv_dict.get('fassade_nord', 0)
fassade_ost = pv_dict.get('fassade_ost', 0)
carport = pv_dict.get('carport', 0)
PV_neu = dach_süd + dach_ostwest + fassade_süd + fassade_west + fassade_nord + fassade_ost + carport
PV_Dach = dach_süd + dach_ostwest + carport
PV_Fassade = fassade_süd + fassade_west + fassade_nord + fassade_ost
capex_pv_fnct = eval(pv_dict.get('capex_pv_fnct', 0))
Capex_PV = pv_dict.get('Capex_PV', 0)
A_Bedarf = pv_dict.get('A_Bedarf', 0)
BK_PV = pv_dict.get('BK_PV', 0)
EIV_pv_fnct = eval(pv_dict.get('EIV_pv_fnct', 0))
EIV = pv_dict.get('EIV', 0)
Vergleich_Fassade = pv_dict.get('Vergleich_Fassade', 0)

"Speicher aus excel einlesen"
df_sp = pd.read_excel('input_parameter.xlsx',5)
variable_names = df_sp['Variabel'].tolist()
variable_values = df_sp['Wert'].tolist()
sp_dict = dict(zip(variable_names, variable_values))
SP_N = int(sp_dict.get('SP_N', 0))
SP_Kap = int(sp_dict.get('SP_Kap', 0))
if SP_Sim == True:
    Speicher = list(range(0, SP_N*SP_Kap, SP_Kap))
else:
    Speicher = list(range(0, 1, 1))
Eigenverbrauch = sp_dict.get('Eigenverbrauch', 0)
Faktor_Grenze = sp_dict.get('Faktor_Grenze', 0)
Faktor_laden = sp_dict.get('Faktor_laden', 0)
ladeverlust = sp_dict.get('ladeverlust', 0)
Eigenverbrauch_grenze = sp_dict.get('Eigenverbrauch_grenze', 0)
capex_sp_fnct = eval(sp_dict.get('capex_sp_fnct', 0))
Capex_SP = sp_dict.get('Capex_SP', 0)
BK_SP = sp_dict.get('BK_SP', 0)
Faktor_vergünstigung = sp_dict.get('Faktor_vergünstigung', 0)

"ELKW aus excel einlesen"
df_elkw = pd.read_excel('input_parameter.xlsx',6)
variable_names = df_elkw['Variabel'].tolist()
variable_values = df_elkw['Wert'].tolist()
elkw_dict = dict(zip(variable_names, variable_values))
elkw_bat_pak = elkw_dict.get('elkw_batteriepaket', 0)
elkw_S = elkw_dict.get('elkw_S', 0)
elkw1_S = elkw_S * elkw_bat_pak
elkw2_S = elkw_S * elkw_bat_pak
elkw3_S = elkw_S * elkw_bat_pak
verbrauch_elkw = elkw_dict.get('verbrauch_elkw', 0)
ladeleistungDC = elkw_dict.get('ladeleistungDC', 0)
schnellladung = elkw_dict.get('schnellladung', 0)
elkw_reserve_akku = elkw_dict.get('elkw_reserve_akku', 0)
elkw_soc_limit_ruhezeit = elkw_dict.get('elkw_soc_limit_ruhezeit', 0)
grenze_soc_raststätte_laden = elkw_dict.get('grenze_soc_raststätte_laden', 0)
laden_30min_raststätte = schnellladung/2 #kWh mit 100 kW laden = 50 kWh

"WELKW aus excel einlesen"
df_welkw = pd.read_excel('input_parameter.xlsx',7)
variable_names = df_welkw['Variabel'].tolist()
variable_values = df_welkw['Wert'].tolist()
welkw_dict = dict(zip(variable_names, variable_values))
kapazität_wechselbatterie = welkw_dict.get('kapazität_wechselbatterie', 0)
welkw_S = welkw_dict.get('welkw_S', 0)
welkw1_S = welkw_S * kapazität_wechselbatterie
welkw2_S = welkw_S * kapazität_wechselbatterie
welkw3_S = welkw_S * kapazität_wechselbatterie
verbrauch_welkw = welkw_dict.get('verbrauch_welkw', 0)
welkw_reserve_akku = welkw_dict.get('welkw_reserve_akku', 0)
welkw_soc_limit_ruhezeit = welkw_dict.get('welkw_soc_limit_ruhezeit', 0)
soc_fuer_wechsel = welkw_dict.get('soc_fuer_wechsel', 0)

"EPKW aus excel einlesen"
df_epkw = pd.read_excel('input_parameter.xlsx',8)
variable_names = df_epkw['Variabel'].tolist()
variable_values = df_epkw['Wert'].tolist()
epkw_dict = dict(zip(variable_names, variable_values))
N_Ladestationen = epkw_dict.get('N_Ladestationen', 0)
durchsch_kapazität = epkw_dict.get('durchsch_kapazität', 0)
epkw_ladekapazitaet = epkw_dict.get('epkw_ladekapazitaet', 0)
ladeleistung_stationen = epkw_dict.get('ladeleistung_stationen', 0)
reduktion_ladeleistung = epkw_dict.get('reduktion_ladeleistung', 0)
epkw_wochenstart_kWh = epkw_dict.get('epkw_wochenstart_kWh', 0)