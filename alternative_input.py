import os
import sys
sys.path.insert(1, 'Module')
import pandas as pd
import numpy as np
import time
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime
from tkinter import *

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
