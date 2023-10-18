import pandas as pd
from pathlib import Path
import os
import glob
import time

""" Provide Excel extraction funcitons for Weekly dashboard code
"""


def get_file():
    time.sleep(4)

    # * means all if need specific format then *.csv
    downloads_path = str(Path.home() / "Downloads/*")

    list_of_files = glob.glob(downloads_path)
    latest_file = max(list_of_files, key=os.path.getctime)

    return latest_file


def get_attendance_COMPASS():
    file = get_file()
    df = pd.read_excel(file,
                       skiprows=41,
                       usecols=['Participant', 'Days Attended'])
    df = df.dropna()
    df = df[df['Days Attended'] != 0]
    os.remove(file)
    return len(df)


def get_attendance_Beacon(tab):
    file = get_file()
    df = pd.read_excel(file,
                       tab,
                       skiprows=7,
                       usecols=['Participant', 'Actual Hours'])
    df = df.dropna()
    df = df[df['Actual Hours'] != '0']
    if tab == 'Sheet4':
        os.remove(file)
    return len(df)


def get_attendance_aLit():
    file = get_file()
    df = pd.read_excel(file,
                       skiprows=41,
                       usecols=['Participant', 'Actual Hours '])
    df = df.dropna()
    df = df[df['Actual Hours '] != '00:00']
    os.remove(file)
    return len(df)  # , df['Actual Hours '].sum()


def get_enrollment_Beacon():
    file = get_file()
    df = pd.read_excel(file,
                       skiprows=25,
                       usecols=['Workscope', 'Total Enrollment'])
    df = df.dropna()
    os.remove(file)
    return df['Total Enrollment'].iloc[0]


def get_ROP_CES(date):
    file = get_file()
    df = pd.read_excel(file,
                       skiprows=24,
                       usecols=['Date (Monday)', 'ROP Weekly Average %'])
    df = df.dropna()
    df = df[df['Date (Monday)'] == date]
    os.remove(file)
    return df['ROP Weekly Average %'].iloc[0]


def get_ROP_CMS(date):
    file = get_file()
    df = pd.read_excel(file,
                       skiprows=22,
                       usecols=['Date (Monday)', 'Cumulative ROP (%)'])
    df = df.dropna()
    df = df[df['Date (Monday)'] == date]
    os.remove(file)
    return df['Cumulative ROP (%)'].iloc[0]


def get_ROP_CYEPe(date):
    file = get_file()
    df = pd.read_excel(file, skiprows=25)
    df.columns = [x.replace("\n", "") for x in df.columns.to_list()]
    df_1 = df[['Date(Monday)', 'Cumulative ROP (%)']]
    df_2 = df[['Date(Monday).1', 'Cumulative ROP(%)']]
    df_2 = df_2.rename(columns={'Date(Monday).1': 'Date(Monday)',
                                'Cumulative ROP(%)': 'Cumulative ROP (%)'})

    df = pd.concat([df_1, df_2])
    df = df.dropna()
    df = df[df['Date(Monday)'] == date]
    os.remove(file)

    return df['Cumulative ROP (%)'].iloc[0]


def get_ROP_CYEPhs(date):
    file = get_file()
    df = pd.read_excel(file, skiprows=25)
    df.columns = [x.replace("\n", " ") for x in df.columns.to_list()]
    df_1 = df[['Date (Monday)', 'Cumulative ROP (%)']]
    df_2 = df[['Date(Monday)', 'Cumulative  ROP (%)']]
    df_2 = df_2.rename(columns={'Date(Monday)': 'Date (Monday)',
                                'Cumulative  ROP (%)': 'Cumulative ROP (%)'})

    df = pd.concat([df_1, df_2])
    df = df.dropna()
    df = df[df['Date (Monday)'] == date]
    os.remove(file)

    return df['Cumulative ROP (%)'].iloc[0]


def get_ROP_B(sheet, date):
    file = get_file()
    df = pd.read_excel(file,
                       sheet,
                       skiprows=7,
                       usecols=['Date (Monday)', 'Rop % (Begins Week  10)'])
    date = pd.to_datetime(date)
    df = df.dropna()
    df = df[df['Date (Monday)'] == date]

    if sheet == 'Sheet3':
        os.remove(file)

    return df['Rop % (Begins Week  10)'].iloc[0]
