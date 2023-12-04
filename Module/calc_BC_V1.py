import numpy as np
from read_input_excel import EIV_pv_fnct,kapazität_wechselbatterie,tarif_schnellladen,Tar_var,r_LKW1,r_LKW2,r_LKW3,Zins_WACC,PV_Fassade,PV_Dach,fassade_süd,fassade_west,fassade_nord,fassade_ost,N_Ladestationen,A_Bedarf,capex_pv_fnct

#Daten für Öl Vergleich
öl_verbrauch = 900000  # kwh
öl_tarif = 0.13 * Tar_var  # CHF/kWh

#best_LKW
lkw_inv =  590812.35
lkw_betriebskosten =   36470.40
lkw_Energiekosten = 36684.02 /1.93*(1.93-0.8)* Tar_var
lkw_steuern =  36684.02 - lkw_Energiekosten
lkw_treibstoff =  lkw_Energiekosten + lkw_steuern
lkw_strecke = r_LKW1["Km gefahren"].sum() + r_LKW2["Km gefahren"].sum() + r_LKW3["Km gefahren"].sum()
N_lkw = 20 #Lebensdauer 800'000km max, 40'000 km pro jahr
Z_lkw = Zins_WACC #Zinssatz
A_lkw = (Z_lkw*(1+Z_lkw)**N_lkw)/((1+Z_lkw)**N_lkw-1) #Annuität
JahreskostenLKW_var = lkw_inv*(A_lkw)+lkw_betriebskosten+lkw_treibstoff
jährliche_Ausgaben_LKW = lkw_betriebskosten + lkw_treibstoff +lkw_inv * (Z_lkw / 2)

# ELKW
elkw_inv_batt = 15000 * 3 #Angabe bax
elkw_inv =   190000 - elkw_inv_batt #Totale Kosten abzüglich batterie = rest vom fahrzeug (angabe bax)
elkw_betriebskosten = 550*12 #Wartungsverrtag bsp Renault, ein drittel abziehen
N_elkw = 20 #Lebensdauer 800'000km max, 40'000 km pro jahr
N_batt = 10
Z_elkw = Zins_WACC #Zinssatz
A_elkw = (Z_elkw*(1+Z_elkw)**N_elkw)/((1+Z_elkw)**N_elkw-1) #Annuität
A_elkw_batt = (Z_elkw*(1+Z_elkw)**N_batt)/((1+Z_elkw)**N_batt-1)
JahreskostenELKW_var = elkw_inv*A_elkw + elkw_betriebskosten + elkw_inv_batt*A_elkw_batt

dc_ladestation_inv = 50000 + 10000 + 700
dc_ladestation_unterhalt = 1000 + 14.90 *12
N_dc_lad = 13
A_dc_ladestation = (Z_elkw*(1+Z_elkw)**N_dc_lad)/((1+Z_elkw)**N_dc_lad-1) #lebensdauer wie batterie 10 jahre
Jahreskosten_dc_ladestation = (dc_ladestation_inv*A_dc_ladestation + dc_ladestation_unterhalt)

# WELKW
welkw_inv = 190000 - elkw_inv_batt #Totale Kosten abzüglich batterie = rest vom fahrzeug // Gleich wie bei elkw
Kosten_Wechselbatterie = 20 + kapazität_wechselbatterie * tarif_schnellladen # Grundgebühr plus energiekosten kwh CHF
welkw_betriebskosten = 550*12 #Wartungsverrtag bsp Renault
welkw_miete_wechselbatterie = (elkw_inv_batt/3/10)*1.5 * 3 #aufrundung der batteriekosten elkw auf jahr gerechnet // multipliziert mit anzahl
welkw_miete_wechselbatterie = 300*1.26*12 #neue Annahme von Nio, 300 CHF/100kWh/m
A_elkw = (Z_elkw*(1+Z_elkw)**N_elkw)/((1+Z_elkw)**N_elkw-1) #Annuität
JahreskostenWELKW_var = welkw_inv*A_elkw + welkw_betriebskosten + welkw_miete_wechselbatterie

ac_ladestation_inv = (1200 + 1000 + 700) *N_Ladestationen
ac_ladestation_unterhalt = (500 + 14.90 *12) * N_Ladestationen
N_ac_lad = 9
A_ac_ladestation = (Z_elkw*(1+Z_elkw)**N_ac_lad)/((1+Z_elkw)**N_ac_lad-1) #lebensdauer wie batterie 10 jahre
Jahreskosten_ac_ladestation = ac_ladestation_inv*A_ac_ladestation + ac_ladestation_unterhalt

#Business-Case PV
Vergleich_Fassade = 70 # CHF/m2
N_PV = 25 #Lebensdauer
Z_PV = Zins_WACC #Zinssatz
A_PV = (Z_PV*(1+Z_PV)**N_PV)/((1+Z_PV)**N_PV-1) #Annuität
BK_PV =  0.01 #Betriebskosten % Invkosten
Faktor_Fassade = 1.5 #Mehrkosten für Installation an Fassade (Montage, Verkabelung) nach Axpo 50%
"Anpassen der Funktion zur Berechnung der Investitionskosten aus der Marktbeobachtungsstudie (Stand 2020) + Aktualisieren im File Input parameter"
if capex_pv_fnct:
    Capex_PV = ((5523/(PV_Dach)**0.4862)+156.2*np.exp(-0.2321*(PV_Dach))+578.4)*(PV_Dach)  #Investitionskosten nach Marktbeobachtungsstudie 2020
if PV_Fassade > 0 and capex_pv_fnct:
    Capex_PV = ((5523 / (PV_Dach) ** 0.4862) + 156.2 * np.exp(-0.2321 * (PV_Dach)) + 578.4) * (PV_Dach) + ((5523 / (PV_Fassade) ** 0.4862) + 156.2 * np.exp(-0.2321 * (PV_Fassade)) + 578.4) * (PV_Fassade) * Faktor_Fassade
    Einsparung_Blech = (fassade_süd+ fassade_west+ fassade_nord + fassade_ost)*A_Bedarf * Vergleich_Fassade
    Capex_PV = Capex_PV - Einsparung_Blech
"Anpassen der Funktion zur Berechnung der EIV (Stand 2023) + Aktualisieren im File Input parameter"
if EIV_pv_fnct:
    EIV_calc = (-0.0019*PV_Dach**2+274.94*PV_Dach+3630) + (-0.0019*PV_Fassade**2+374.94*PV_Fassade+3630)
else:
    from read_input_excel import EIV
    EIV_calc = EIV
JahreskostenPV_var = Capex_PV*(A_PV+BK_PV)
jährliche_Ausgaben_PV = Capex_PV * (BK_PV + Z_PV / 2)

#Business-Case Speicher
N_BS = 15
Z_BS = Zins_WACC
A_BS = (Z_BS*(1+Z_BS)**N_BS)/((1+Z_BS)**N_BS-1)
"Da der Capex des Batteriespeicher abhängig von der Grösse ist wird dieser in der Speicher Schlaufe berechnet, Suche nach Capex_SP"