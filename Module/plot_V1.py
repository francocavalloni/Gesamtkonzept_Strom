import pandas as pd
import numpy as np
import time
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime

def plt_elkw(ELKW_Sim,INFO_ELKW,data_S):
    for elkw in range(1, ELKW_Sim + 1):
        if elkw == 1:
            time_windows = [('2022-04-27 00:00:00', '2022-04-28 00.00:00'),
                            ('2022-05-23 00:00:00', '2022-05-30 00:00:00'),
                            ('2022-11-14 00:00:00', '2022-11-21 00:00:00')]
        else:
            time_windows = [('2022-05-23 00:00:00', '2022-05-28 00:00:00'),
                            ('2022-11-14 00:00:00', '2022-11-21 00:00:00')]
        for start_time, end_time in time_windows:
            subset_lkw = INFO_ELKW.loc[(INFO_ELKW['Time'] >= start_time) & (INFO_ELKW['Time'] <= end_time)]
            subset_lg = data_S.loc[(INFO_ELKW['Time'] >= start_time) & (INFO_ELKW['Time'] <= end_time)]

            sns.set(rc={'figure.dpi': 300})
            sns.set_style("whitegrid")
            fig, ax1 = plt.subplots(figsize=(10, 6))

            sns.lineplot(x="Time", y="kW_PV_S", data=subset_lg, ax=ax1, linewidth=2.5, color='orange',
                         label="Last ohne E-LKW")
            sns.lineplot(x="Time", y="kW_RS_S", data=subset_lg, ax=ax1, linewidth=2.5, color='gold',
                         label="Rücksp ohne E-LKW")
            sns.lineplot(x="Time", y="Last_ELKW" + str(elkw), data=subset_lkw, ax=ax1, linewidth=2.5,
                         color='green',
                         label="Last mit E-LKW")
            sns.lineplot(x="Time", y="Ruecksp_ELKW" + str(elkw), data=subset_lkw, ax=ax1, linewidth=2.5,
                         color='yellowgreen',
                         label="Rücksp mit E-LKW")

            ax1.set_ylim(0, 1000)
            ax1.set_ylabel('Lastgang [kW]')

            ax2 = ax1.twinx()
            sns.lineplot(x="Time", y="ELKW" + str(elkw) + "_SOC", data=subset_lkw, ax=ax2, linewidth=2.5,
                         color='deepskyblue', label="SOC")
            ax2.set_ylim(0, 1)
            ax2.set_ylabel('SOC')

            # Set the x-axis limits to match the start and end time
            ax1.set_xlim(pd.Timestamp(start_time), pd.Timestamp(end_time))

            # Set the major ticks to be at every 12 hours and format the tick labels
            hours = mdates.HourLocator(interval=12)
            fmt = mdates.DateFormatter('%H:%M\n%d.%m')
            ax1.xaxis.set_major_locator(hours)
            ax1.xaxis.set_major_formatter(fmt)

            # Set the first tick to be at the start time
            ax1.xaxis.set_tick_params(which='major', pad=15)
            first_tick = pd.Timestamp(start_time).floor('D')
            ax1.set_xticks(pd.date_range(first_tick, pd.Timestamp(end_time), freq='12H'))

            ax1.legend_.remove()
            ax2.legend_.remove()
            fig.legend(loc="upper right", bbox_to_anchor=(1, 1), bbox_transform=ax2.transAxes)
            plt.title('Auswirkung E-LKW' + str(elkw))
            plt.show()
    return plt.show()

def plt_welkw(WELKW_Sim,INFO_WELKW,data_S):
    for welkw in range(1, WELKW_Sim + 1):
        if welkw == 1:
            time_windows = [('2022-04-27 00:00:00', '2022-04-28 00.00:00'),
                            ('2022-05-23 00:00:00', '2022-05-30 00:00:00'),
                            ('2022-11-14 00:00:00', '2022-11-21 00:00:00')]
        else:
            time_windows = [('2022-05-23 00:00:00', '2022-05-30 00:00:00'),
                            ('2022-11-14 00:00:00', '2022-11-21 00:00:00')]
        for start_time, end_time in time_windows:
            subset_lkw = INFO_WELKW.loc[(INFO_WELKW['Time'] >= start_time) & (INFO_WELKW['Time'] <= end_time)]
            subset_lg = data_S.loc[(INFO_WELKW['Time'] >= start_time) & (INFO_WELKW['Time'] <= end_time)]

            sns.set(rc={'figure.dpi': 300})
            sns.set_style("whitegrid")
            fig, ax1 = plt.subplots(figsize=(10, 6))

            sns.lineplot(x="Time", y="kW_PV_S", data=subset_lg, ax=ax1, linewidth=2.5, color='orange',
                         label="Last best.")
            sns.lineplot(x="Time", y="kW_RS_S", data=subset_lg, ax=ax1, linewidth=2.5, color='gold',
                         label="Rücksp best.")
            sns.lineplot(x="Time", y="Last_WELKW" + str(welkw), data=subset_lkw, ax=ax1, linewidth=2.5,
                         color='green',
                         label="Last WELKW")
            sns.lineplot(x="Time", y="Ruecksp_WELKW" + str(welkw), data=subset_lkw, ax=ax1, linewidth=2.5,
                         color='yellowgreen',
                         label="Rücksp WELKW")
            ax1.set_ylim(0, 1000)
            ax1.set_ylabel('Lastgang [kW]')

            ax2 = ax1.twinx()
            sns.lineplot(x="Time", y="WELKW" + str(welkw) + "_SOC", data=subset_lkw, ax=ax2, linewidth=2.5,
                         color='deepskyblue', label="SOC")
            ax2.set_ylim(0, 1)
            ax2.set_ylabel('SOC')

            # Set the x-axis limits to match the start and end time
            ax1.set_xlim(pd.Timestamp(start_time), pd.Timestamp(end_time))

            # Set the major ticks to be at every 12 hours and format the tick labels
            hours = mdates.HourLocator(interval=12)
            fmt = mdates.DateFormatter('%H:%M\n%d.%m')
            ax1.xaxis.set_major_locator(hours)
            ax1.xaxis.set_major_formatter(fmt)

            # Set the first tick to be at the start time
            ax1.xaxis.set_tick_params(which='major', pad=15)
            first_tick = pd.Timestamp(start_time).floor('D')
            ax1.set_xticks(pd.date_range(first_tick, pd.Timestamp(end_time), freq='12H'))

            ax1.legend_.remove()
            ax2.legend_.remove()
            fig.legend(loc="upper right", bbox_to_anchor=(1, 1), bbox_transform=ax2.transAxes)
            plt.title('Auswirkung WE-LKW' + str(welkw))
    return plt.show()

def plt_epkw(INFO_EPKW,data_S,ELKW_Sim,INFO_ELKW,WELKW_Sim, INFO_WELKW):
    elkw = ELKW_Sim
    welkw = WELKW_Sim
    time_windows = [('2022-05-23 00:00:00', '2022-05-30 00:00:00'),
                    ('2022-11-14 00:00:00', '2022-11-21 00:00:00')]
    for start_time, end_time in time_windows:
        subset_epkw = INFO_EPKW.loc[(INFO_EPKW['Time'] >= start_time) & (INFO_EPKW['Time'] <= end_time)]

        if ELKW_Sim >=1:
            subset_lkw = INFO_ELKW.loc[(INFO_ELKW['Time'] >= start_time) & (INFO_ELKW['Time'] <= end_time)]
        elif WELKW_Sim >=1:
            subset_lkw = INFO_WELKW.loc[(INFO_WELKW['Time'] >= start_time) & (INFO_WELKW['Time'] <= end_time)]
        else:
            subset_lg = data_S.loc[(INFO_EPKW['Time'] >= start_time) & (INFO_EPKW['Time'] <= end_time)]

        sns.set(rc={'figure.dpi': 300})
        sns.set_style("whitegrid")
        fig, ax1 = plt.subplots(figsize=(10, 6))

        if ELKW_Sim >=1:
            sns.lineplot(x="Time", y="Last_ELKW" + str(elkw), data=subset_lkw, ax=ax1, linewidth=2.5,
                         color='orange',
                         label="Last ohne EPKW")
            sns.lineplot(x="Time", y="Ruecksp_ELKW" + str(elkw), data=subset_lkw, ax=ax1, linewidth=2.5,
                         color='gold',
                         label="Rücksp ohne EPKW")
        elif WELKW_Sim >=1:
            sns.lineplot(x="Time", y="Last_WELKW" + str(welkw), data=subset_lkw, ax=ax1, linewidth=2.5,
                         color='orange',
                         label="Last ohne EPKW")
            sns.lineplot(x="Time", y="Ruecksp_WELKW" + str(welkw), data=subset_lkw, ax=ax1, linewidth=2.5,
                         color='gold',
                         label="Rücksp ohne EPKW")
        else:
            sns.lineplot(x="Time", y="kW_PV_S", data=subset_lg, ax=ax1, linewidth=2.5, color='orange',
                         label="Last ohne EPKW")
            sns.lineplot(x="Time", y="kW_RS_S", data=subset_lg, ax=ax1, linewidth=2.5, color='gold',
                         label="Rücksp ohne EPKW")

        sns.lineplot(x="Time", y="Last_EPKW", data=subset_epkw, ax=ax1, linewidth=2.5, color='green',
                     label="Last mit EPKW")
        sns.lineplot(x="Time", y="Ruecksp_EPKW", data=subset_epkw, ax=ax1, linewidth=2.5, color='yellowgreen',
                     label="Rücksp mit EPKW")

        ax1.set_ylim(0, 1000)
        ax1.set_ylabel('Lastgang [kW]')

        ax2 = ax1.twinx()
        sns.lineplot(x="Time", y="EPKW_SOC", data=subset_epkw, ax=ax2, linewidth=2.5, color='deepskyblue',
                     label="SOC")
        ax2.set_ylim(0, 1)
        ax2.set_ylabel('SOC')

        # Set the x-axis limits to match the start and end time
        ax1.set_xlim(pd.Timestamp(start_time), pd.Timestamp(end_time))

        # Set the major ticks to be at every 12 hours and format the tick labels
        hours = mdates.HourLocator(interval=12)
        fmt = mdates.DateFormatter('%H:%M\n%d.%m')
        ax1.xaxis.set_major_locator(hours)
        ax1.xaxis.set_major_formatter(fmt)

        # Set the first tick to be at the start time
        ax1.xaxis.set_tick_params(which='major', pad=15)
        first_tick = pd.Timestamp(start_time).floor('D')
        ax1.set_xticks(pd.date_range(first_tick, pd.Timestamp(end_time), freq='12H'))
        ax1.legend_.remove()
        ax2.legend_.remove()
        fig.legend(loc="upper right", bbox_to_anchor=(1, 1), bbox_transform=ax1.transAxes)
        plt.title('Auswirkung E-PKW')
    return plt.show()

def plt_BS(Speicher,PROFIL_SPEICHER,INFO_SPEICHER):
    for S in Speicher:
        if S > 0:
            time_windows = [('2022-07-11 00:00:00', '2022-07-12 00:00:00'),
                            ('2022-01-10 00:00:00', '2022-01-11 00:00:00'),
                            ('2022-03-02 00:00:00', '2022-03-04 00:00:00'),
                            ('2022-09-01 00:00:00', '2022-09-02 00:00:00')]
            for start_time, end_time in time_windows:
                subset_Speicher = PROFIL_SPEICHER.loc[
                    (PROFIL_SPEICHER['Time'] >= start_time) & (PROFIL_SPEICHER['Time'] <= end_time)]
                subset_SOC = INFO_SPEICHER.loc[
                    (INFO_SPEICHER['Time'] >= start_time) & (INFO_SPEICHER['Time'] <= end_time)]

                sns.set(rc={'figure.dpi': 300})
                sns.set_style("whitegrid")
                fig, ax1 = plt.subplots(figsize=(10, 6))

                sns.lineplot(x="Time", y="Last_0", data=subset_Speicher, ax=ax1, linewidth=2.5, color='orange',
                             label="Last ohne Speicher")
                sns.lineplot(x="Time", y="Ruecksp_0", data=subset_Speicher, ax=ax1, linewidth=2.5, color='gold',
                             label="Rücksp ohne Speicher")
                sns.lineplot(x="Time", y="Last_" + str(S), data=subset_Speicher, ax=ax1, linewidth=2.5,
                             color='green',
                             label="Last mit Speicher")
                sns.lineplot(x="Time", y="Ruecksp_" + str(S), data=subset_Speicher, ax=ax1, linewidth=2.5,
                             color='yellowgreen',
                             label="Rücksp mit Speicher")
                max_y_ax = subset_Speicher["Last_0"].max()
                if max_y_ax < 500:
                    y_ax_limit = 500
                elif max_y_ax < 1000:
                    y_ax_limit = 1000
                elif max_y_ax < 1500:
                    y_ax_limit = 1500
                else:
                    y_ax_limit = 2000

                ax1.set_ylim(0, y_ax_limit)
                ax1.set_ylabel('Lastgang [kW]')

                ax2 = ax1.twinx()
                sns.lineplot(x="Time", y=str(S) + "_SOC", data=subset_SOC, ax=ax2, linewidth=2.5,
                             color='deepskyblue', label="SOC")
                ax2.set_ylim(0, 1)
                ax2.set_ylabel('SOC')

                # Set the x-axis limits to match the start and end time
                ax1.set_xlim(pd.Timestamp(start_time), pd.Timestamp(end_time))

                # Set the major ticks to be at every 12 hours and format the tick labels
                hours = mdates.HourLocator(interval=12)
                fmt = mdates.DateFormatter('%H:%M\n%d.%m')
                ax1.xaxis.set_major_locator(hours)
                ax1.xaxis.set_major_formatter(fmt)

                # Set the first tick to be at the start time
                ax1.xaxis.set_tick_params(which='major', pad=15)
                first_tick = pd.Timestamp(start_time).floor('D')
                ax1.set_xticks(pd.date_range(first_tick, pd.Timestamp(end_time), freq='12H'))

                ax1.legend_.remove()
                ax2.legend_.remove()
                fig.legend(loc="upper right", bbox_to_anchor=(1, 1), bbox_transform=ax2.transAxes)
                plt.title('Auswirkung Batteriespeicher' + str(S))
    return plt.show()

def plt_PV(data_S,PV_neu):
    time_windows = [('2022-05-23 00:00:00', '2022-05-30 00:00:00'),
                    ('2022-07-18 00:00:00', '2022-07-25 00:00:00'),
                    ('2022-08-08 00:00:00', '2022-08-15 00:00:00'),
                    ('2022-12-12 00:00:00', '2022-12-19 00:00:00')]
    for start_time, end_time in time_windows:
        subset_lg = data_S.loc[(data_S['Time'] >= start_time) & (data_S['Time'] <= end_time)]

        sns.set(rc={'figure.dpi': 300})
        sns.set_style("whitegrid")
        fig, ax1 = plt.subplots(figsize=(10, 6))

        sns.lineplot(x="Time", y="kW_Last", data=subset_lg, ax=ax1, linewidth=2.5, color='orange',
                     label="Last ohne PV")
        sns.lineplot(x="Time", y="kW_PV_S", data=subset_lg, ax=ax1, linewidth=2.5, color='green',
                     label="Last mit PV")
        sns.lineplot(x="Time", y="kW_RS_S", data=subset_lg, ax=ax1, linewidth=2.5, color='yellowgreen',
                     label="Rücksp mit PV")
        ax1.set_ylim(0, 1000)
        ax1.set_ylabel('Lastgang [kW]')

        # Set the x-axis limits to match the start and end time
        ax1.set_xlim(pd.Timestamp(start_time), pd.Timestamp(end_time))

        # Set the major ticks to be at every 12 hours and format the tick labels
        hours = mdates.HourLocator(interval=12)
        fmt = mdates.DateFormatter('%H:%M\n%d.%m')
        ax1.xaxis.set_major_locator(hours)
        ax1.xaxis.set_major_formatter(fmt)

        # Set the first tick to be at the start time
        ax1.xaxis.set_tick_params(which='major', pad=15)
        first_tick = pd.Timestamp(start_time).floor('D')
        ax1.set_xticks(pd.date_range(first_tick, pd.Timestamp(end_time), freq='12H'))
        ax1.legend_.remove()
        fig.legend(loc="upper right", bbox_to_anchor=(1, 1), bbox_transform=ax1.transAxes)
        plt.title('Auswirkung PV ' + str(PV_neu) + 'kWp')
    return plt.show()