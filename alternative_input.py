import pandas as pd
import sys
import os
sys.path.insert(1, 'Module')
current_directory = os.getcwd()

"Inputs für Variantenbildung"
LG_Sim = True
PV_Sim = True
ELKW_Sim = 3
WELKW_Sim = 0
EPKW_Sim = False
SP_Sim = True
excel = True
plot = False
LKW_Sim = False

"Paths einlesen"
output_excel_S = "Test_Output.xlsx"  #Ausgabedatei
r_DIR = "Daten_Input"
page_sim_LG = 3
page_an_LG = 3
r_LG_S = pd.read_excel(os.path.join(current_directory,r_DIR,"Beispiel_Lastgang_einlesen.xlsx"), page_sim_LG) #Lastgang 1
if page_an_LG == page_sim_LG:
    r_LG_A = r_LG_S
else:
    r_LG_A = pd.read_excel(os.path.join(current_directory,r_DIR,"Beispiel_Lastgang_einlesen.xlsx"), page_an_LG)
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

"##### Input Parameter #####"
output_excel_S = "Variante_2b_ohne_öl.xlsx"  #anpassen nach Daten einlesen !!einsetzen!! je nach Variante Lastgang anpassen unten Zeile 47!
Sim_LG = 1 #Bilanzierung des Lastgangs [ 0 = aus / 1 = ein ]
ELKW_Sim = 3 #anzahl elkw 3 , [ 0 = aus ]
WELKW_Sim = 0 #anzahl welkw , [ 0 = aus]
EPKW_Sim = 1 # [ 0 = aus / 1 = ein ]
Speicher = list(range(0, 300, 136)) # (min, max, schrittgrösse) [ max < schrittgrösse -> nur PV bilanzierung ]
# Eingabe der PV-Anlagen in kWp
dach_süd = 0
dach_ostwest = 1206
fassade_süd = 299
fassade_west = 213
fassade_nord = 299
fassade_ost = 231
carport = 300
excel = False
plot = True

"##### Daten einlesen #####"
# Die Daten für Lastgang und PV-Anlage werden in einem Format von 15 Min Werten über ein ganzes Jahr eingelesen
Feiertage = ["01-01","01-02", "12-25", "12-26"] #Tage ohne Arbeitsbetrieb für Eigenverbrauch anstatt Peak-Shaving
# Daten Mobilität, Format für einlesen
current_directory = os.getcwd()
r_LKW1 = pd.read_excel(os.path.join(current_directory, 'Daten_Lupfig', 'Mobiltät_LKW_2022_Fahrdaten.xlsx'), 0)
r_LKW2 = pd.read_excel(os.path.join(current_directory, 'Daten_Lupfig', 'Mobiltät_LKW_2022_Fahrdaten.xlsx'), 1)
r_LKW3 = pd.read_excel(os.path.join(current_directory, 'Daten_Lupfig', 'Mobiltät_LKW_2022_Fahrdaten.xlsx'), 2)
r_LG = pd.read_excel(os.path.join(current_directory, 'Daten_Lupfig', 'Analyse_Lastgang22_Lupfig_Taracell_20230315.xlsx'), 4) #Lastgang 2
LG_name = "Effizient"
####### !!einsetzen für r_LG!!: r_LG_Trend_Künten (var4), r_LG_Effizient (var5) oder r_LG_öl (Ref, var1,2,3) ######

"einlesen der PV-Daten mit Polysun simuliert oder Messdaten"
r_PV_F = pd.read_excel(os.path.join(current_directory, 'Daten_Lupfig', 'Berechnung_und_PolysunOutput_20221219.xlsx'), 2) #PV_Fassaden
r_PV_DS = pd.read_excel(os.path.join(current_directory, 'Daten_Lupfig', 'Berechnung_und_PolysunOutput_20221219.xlsx'), 3) #PV_Dach_Süd
r_PV_DOW = pd.read_excel(os.path.join(current_directory, 'Daten_Lupfig', 'Berechnung_und_PolysunOutput_20221219.xlsx'), 4) #PV_Dach_OstWest
r_PV_CP = pd.read_excel(os.path.join(current_directory, 'Daten_Lupfig', 'Berechnung_und_PolysunOutput_20221219.xlsx'), 5) #PV_Dach_carport
spalte_analyse = "ANALYSE_SPITZE"


"##### TARIFE #####"
# Faktor Tarifvariation Energieanteil Arbeitspreis: 1 = Standard
Tar_var = 1.15
# Faktor Tarifvariation Leistungspreis: 1 = Standard
Lei_var = 1.15
leistungsspitzenpreis = 7.5 * Lei_var # CHF/kW
hochtarif = (0.1079 * Tar_var +0.0204 + 0.0296)*1.077 # CHF/kWh
niedertarif = (0.0964 * Tar_var +0.014 + 0.0296)*1.077 #0.16927 * Tar_var # CHF/kWh
Referenzmarktpreis = 76.92 # Q1 2023
rueckspeisungstarif = ((Referenzmarktpreis/1000-0.003)* Tar_var) + 0.0200 # CHF/kWhN_BS
tarif_schnellladen = 0.5 * Tar_var #CHF/kWh
öl_verbrauch = 900000 #kwh
öl_tarif = 0.13 * Tar_var #CHF/kWh
Zins_WACC = 0.0413 #2024
#Zeiten für Hochtarif
zeit_hochtarif_woche_start = datetime.datetime.strptime("06:00:00", "%H:%M:%S")
zeit_hochtarif_woche_ende = datetime.datetime.strptime("20:00:00", "%H:%M:%S")
zeit_hochtarif_samstag_start = datetime.datetime.strptime("06:00:00", "%H:%M:%S")
zeit_hochtarif_sonntag_ende = datetime.datetime.strptime("13:00:00", "%H:%M:%S")

# Dimension PV-Anlage
PV_neu = dach_süd + dach_ostwest + fassade_süd + fassade_west + fassade_nord + fassade_ost + carport
PV_Dach = dach_süd + dach_ostwest
PV_Fassade = fassade_süd + fassade_west + fassade_nord + fassade_ost
PV_Carport = carport
output_excel_A = "Analyse_2022_PV_"+str(PV_neu)+"kWp.xlsx"

# ELKW
elkw_bax_batteriepaket = 42 #kwh
elkw1_S = 126 #kWh 3 Pakete
elkw2_S = 126 #kWh 5 Pakete
elkw3_S = 126 #kWh
Verbrauch = 0.63 #kWh/km nach Bax
ladeleistungDC = 90 #kW
schnellladung = 100 #kW an schnellladesäulen
reserve_akku = 0 # 0%
ELKW_SOC_limit_Ruhezeit = 0.1 #Ladelimite von Netz ausserhalb der Betriebszeiten (Status = 0)
Grenze_SOC_Raststätte_laden = 0.05
laden_30min_raststätte = 50 #kWh mit 100 kW laden

# WELKW
Kapazität_Wechselbatterie = 42 #kWh
welkw1_S = 126 #kWh
welkw2_S = 126 #kWh
welkw3_S = 126 #kWh
Kosten_Wechselbatterie = 20 + Kapazität_Wechselbatterie * tarif_schnellladen # Grundgebühr plus energiekosten kwh CHF
SOC_fuer_wechsel = 0.1

# EPKW
N_Ladestationen = 10
durchsch_kapazität = 68 #durchschnittlich nutzbare kapazität epkw
durchsch_Reichweite = 350
durchsch_verbrauch = 0.2 #kWh/km
epkw_ladekapazitaet = 10 #kWh, 50km, wie viel laden pro tag!
ladeleistung_stationen = 22 #kW
reduktion_ladeleistung = 0.25 # über die 2h nur mit 25% laden
epkw_wochenstart_kWh = 0.5

#Speicher Parameter
Eigenverbrauch = 1    ## 0, 1, 2, 3
# 0 nur Peak-Shaving und keine EVerhöhung
# 1 nur wochenende EVerhöhung
# 2 auch tiefe lasten EVerhöhung (eigenverbrauchsgrenze) zb. vor und nach rückspeisung
# 3 nur Eigenverbrauch, heisst kein Netzladen
# Sicherheit laden
Faktor_Grenze = 0.01  # Sicherheitsfaktor (bsp 0.1 = 10% Sicherheit, der angestrebte max. Peak wird um 10% der Lastreduktion erhöht)
# Grenze Speicher laden von Netz im Lastprofil [Faktor für Bezug auf Spitzen Grenze, Speicher abhängig]
Faktor_laden = 1
ladeverlust = 0.1
# Wenn last unter dieser grenze Speicher entladen, kW, sollte so gewählt werden, dass nur vor und nach PV Rückspeisung entladen wird, nicht bei tiefer Grundlast
Eigenverbrauch_grenze = 50
#speicher Kostendegression bsp 0.2 = 20% Kostensenkung
Faktor_vergünstigung = 0
