import pandas as pd
import datetime
import sys
sys.path.insert(1, 'Module')
from read_LG_PV_V1 import data_S
from read_input_excel import ELKW_Sim, WELKW_Sim, r_LKW1, r_LKW2, r_LKW3, elkw1_S, elkw2_S, elkw3_S, welkw1_S, welkw2_S, welkw3_S, verbrauch_welkw, verbrauch_elkw, schnellladung

print("LKW Fahrdaten einlesen")
counter = 1
if ELKW_Sim:
    Verbrauch = verbrauch_elkw
if WELKW_Sim:
    Verbrauch = verbrauch_welkw
for r_LKW in (r_LKW1, r_LKW2, r_LKW3):
    r_LKW['datetime_start'] = None
    r_LKW['month-day'] = None
    index_date = r_LKW.columns.get_loc('Tag')
    index_time = r_LKW.columns.get_loc('LKW unterwegs')
    index_datetime = r_LKW.columns.get_loc('datetime_start')
    for row in range(0, len(r_LKW)):
        dateA = datetime.datetime.strptime(r_LKW.iat[row, index_date], "%d.%m.%Y   %A")
        datum: str = datetime.datetime.date(dateA)
        datum_n: str = datetime.datetime.date(dateA).strftime('%m-%d')
        stime = datetime.datetime.strptime(r_LKW.iat[row, index_time][:5] + ":00", "%H:%M:%S")
        zeit: str = datetime.datetime.time(stime)
        r_LKW.iat[row, r_LKW.columns.get_loc('month-day')] = datum_n
        r_LKW.iat[row, index_datetime] = datetime.datetime.combine(datum, zeit)

    r_LKW["datetime_start"] = r_LKW["datetime_start"].astype('datetime64')
    r_LKW["datetime_start"] = pd.to_datetime(r_LKW["datetime_start"], format="%d.%m.%Y %H:%M:%S")
    r_LKW["datetime_start"] = r_LKW["datetime_start"].apply(lambda x: x + datetime.timedelta(minutes=15 - x.minute % 15))

    # Load the data from DF2
    df2 = r_LKW
    df2 = df2.rename(columns={'datetime_start': 'starting_time'})
    df2 = df2.rename(columns={'Arbeitszeit': 'working_time'})
    df2 = df2.rename(columns={'Lenkzeit': 'driving_time'})
    df2 = df2.rename(columns={'Pause': 'break_time'})
    df2 = df2.rename(columns={'Km gefahren': 'distance'})
    df2 = df2.rename(columns={'Fz beladen in Künten': 'stops'})
    DF2 = df2

    # Create an empty dataframe for DF3
    driving_dist = str(counter) + ' driving_distance'
    status = str(counter) + ' status'
    df3 = pd.DataFrame(columns=['timestamp', driving_dist, status])
    DF3 = df3

    # df for ziel und total distance
    target = str(counter) + ' Ziel'
    tot_distance = str(counter) + ' total_distance'
    beladen = str(counter) + ' Anzahl beladen Lupfig'
    df_z_d = pd.DataFrame(columns=['month-day', target, tot_distance, beladen])
    df_z_d["month-day"] = df2["month-day"]
    df_z_d[tot_distance] = df2["distance"]
    df_z_d[target] = df2["Ziel"]
    df_z_d[beladen] = df2['stops']

    if counter == 1:
        S_ELKW = elkw1_S
    if counter ==2:
        S_ELKW = elkw2_S
    if counter ==3:
        S_ELKW = elkw3_S
    summe_break_diff = 0

    if WELKW_Sim > 0:
        if counter == 1:
            S_ELKW = welkw1_S
        if counter == 2:
            S_ELKW = welkw2_S
        if counter == 3:
            S_ELKW = welkw3_S
        summe_break_diff = 0

    def get_min(time_str):
        """Get seconds from time."""
        h, m = time_str.split(':')
        return int(h) * 60 + int(m)
    def get_sec(time_str):
        """Get seconds from time."""
        h, m, s = time_str.split(':')
        return int(h) * 3600 + int(m) * 60 + int(s)

    # loop through each row in DF2
    for i, row in DF2.iterrows():
        # calculate the start and end timestamps for the current route
        start_time = row['starting_time']
        total_distance = row['distance']
        specific_target = row["Ziel"]
        working_time = round(get_min(row['working_time'])/15)*15
        driving_time = round(get_min(row['driving_time'])/15)*15
        break_time = round(get_min(row['break_time'])/15)*15
        #Pause verlängern, wenn zu kurz für schnelladung
        #TODO Anpassen damit funktioniert, pausen an raststätten sind zu gross:
        if break_time/60*schnellladung < (total_distance*Verbrauch-S_ELKW):
            break_time_alt = break_time
            break_time = round(((total_distance*Verbrauch-S_ELKW)/schnellladung*60)/15)*15
            break_time_diff = break_time - break_time_alt
            summe_break_diff = summe_break_diff + break_time_diff
        end_time = start_time + pd.Timedelta(minutes=driving_time + working_time + break_time)

        # calculate the distance covered in each 15 minute interval
        if driving_time // 15 == 0:
            distance_per_15min = 0
        else:
            distance_per_15min = total_distance / (driving_time // 15)
        stops = row['stops'] * 2
        if stops == 0:
            stops = 1
        half = round(-(-stops//2))/stops
        if stops/2 ==3:
            half = 2/(stops/2)
        # loop through each 15 minute interval between start_time and end_time
        timestamp = start_time
        distance = 0

        while timestamp <= end_time and distance < total_distance:
            # add a row to DF3 for the current timestamp
            new_row = pd.DataFrame([[timestamp, 0, 0]], columns=['timestamp', driving_dist, status])
            DF3 = pd.concat([DF3, new_row])
            # determine the status for the current timestamp
            if timestamp < start_time + pd.Timedelta(minutes=working_time/stops):
                # first working time at Taracell
                DF3.loc[DF3['timestamp'] == timestamp, status] = -1
            elif timestamp < start_time + pd.Timedelta(minutes=working_time/stops + driving_time / stops):
                # first half of driving time
                DF3.loc[DF3['timestamp'] == timestamp, driving_dist] = distance_per_15min
                DF3.loc[DF3['timestamp'] == timestamp, status] = 1
                distance = distance + distance_per_15min
            elif stops >= 4 and timestamp < start_time + pd.Timedelta(minutes=working_time*2/stops+ driving_time / stops):
                DF3.loc[DF3['timestamp'] == timestamp, status] = -3
            elif stops >= 4 and timestamp < start_time + pd.Timedelta(minutes=working_time*2/stops + driving_time*2 / stops):
                DF3.loc[DF3['timestamp'] == timestamp, driving_dist] = distance_per_15min
                DF3.loc[DF3['timestamp'] == timestamp, status] = 1
                distance = distance + distance_per_15min
            elif stops >= 6 and timestamp < start_time + pd.Timedelta(minutes=working_time*3/stops+ driving_time*2 / stops):
                DF3.loc[DF3['timestamp'] == timestamp, status] = -1
            elif stops >= 6 and timestamp < start_time + pd.Timedelta(minutes=working_time*3/stops + driving_time*3 / stops):
                DF3.loc[DF3['timestamp'] == timestamp, driving_dist] = distance_per_15min
                DF3.loc[DF3['timestamp'] == timestamp, status] = 1
                distance = distance + distance_per_15min
            elif stops >= 6 and timestamp < start_time + pd.Timedelta(minutes=working_time*4/stops+ driving_time*3 / stops):
                DF3.loc[DF3['timestamp'] == timestamp, status] = -3
            elif stops >= 6 and timestamp < start_time + pd.Timedelta(minutes=working_time*4/stops + driving_time*4 / stops):
                DF3.loc[DF3['timestamp'] == timestamp, driving_dist] = distance_per_15min
                DF3.loc[DF3['timestamp'] == timestamp, status] = 1
                distance = distance + distance_per_15min
            elif timestamp < start_time + pd.Timedelta(minutes=working_time*half + driving_time*half + break_time):
                # brake time
                DF3.loc[DF3['timestamp'] == timestamp, status] = -2
            elif stops >= 8 and timestamp < end_time - pd.Timedelta(minutes=working_time*3/stops + driving_time*4/stops):
                DF3.loc[DF3['timestamp'] == timestamp, status] = -1
            elif stops >= 8 and timestamp < end_time - pd.Timedelta(minutes=working_time*3/stops + driving_time*3/stops):
                DF3.loc[DF3['timestamp'] == timestamp, driving_dist] = distance_per_15min
                DF3.loc[DF3['timestamp'] == timestamp, status] = 1
                distance = distance + distance_per_15min
            elif stops >= 8 and timestamp < end_time - pd.Timedelta(minutes=working_time*2/stops + driving_time*3/stops):
                DF3.loc[DF3['timestamp'] == timestamp, status] = -3
            elif stops >= 8 and timestamp < end_time - pd.Timedelta(minutes=working_time*2/stops + driving_time*2/stops):
                DF3.loc[DF3['timestamp'] == timestamp, driving_dist] = distance_per_15min
                DF3.loc[DF3['timestamp'] == timestamp, status] = 1
                distance = distance + distance_per_15min
            elif stops >= 4 and timestamp < end_time - pd.Timedelta(minutes=working_time/stops + driving_time*2/stops):
                DF3.loc[DF3['timestamp'] == timestamp, status] = -1
            elif stops >= 4 and timestamp < end_time - pd.Timedelta(minutes=working_time/stops + driving_time/stops):
                DF3.loc[DF3['timestamp'] == timestamp, driving_dist] = distance_per_15min
                DF3.loc[DF3['timestamp'] == timestamp, status] = 1
                distance = distance + distance_per_15min
            elif timestamp < end_time - pd.Timedelta(minutes=driving_time/stops):
                # second half of working time
                DF3.loc[DF3['timestamp'] == timestamp, status] = -3
                if specific_target == "Burnhaupt (FR)":
                    DF3.loc[DF3['timestamp'] == timestamp, status] = -4
            else:
                # second half of driving time
                DF3.loc[DF3['timestamp'] == timestamp, driving_dist] = distance_per_15min
                DF3.loc[DF3['timestamp'] == timestamp, status] = 1
                if timestamp == end_time:
                    DF3.loc[DF3['timestamp'] == timestamp, driving_dist] = total_distance-distance
                distance = distance + distance_per_15min
            # increment timestamp by 15 minutes
            timestamp += pd.Timedelta(minutes=15)

    data_S = pd.merge(data_S, DF3, left_on="Time", right_on="timestamp", how='outer')
    del data_S["timestamp"]
    data_S = pd.merge(data_S, df_z_d, on="month-day", how='left')
    if counter == 1:
        elkw1_summe_break_diff = summe_break_diff
    if counter ==2:
        elkw2_summe_break_diff = summe_break_diff
    if counter ==3:
        elkw3_summe_break_diff = summe_break_diff
    counter = counter +1
data_S = data_S.fillna(0)

"return data_S, elkw1_summe_break_diff, elkw2_summe_break_diff, elkw3_summe_break_diff"