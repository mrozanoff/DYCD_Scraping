import os

import pandas as pd
import PySimpleGUI as sg
from pathlib import Path
from datetime import timedelta, date

# from DYCD_Functions import Webdriver
from DYCD_Functions_01 import UnaccountedAttendanceScraping
from DYCD_ExcelFunctions import get_file

""" Scraping script to pull unaccounted attendance from Dycd connect

TODO:
modules should be lowercase with maybe underscores
classes should be capitalized
function names should be lowercase with underscores
-UPDATE FIND DOWNOADS TO BE LAST ONE, LIKE WEEKLY DASH

take the YD summary from two weeks ago and show the UA col for diffrence?

BLANKS
PS5
CS 61
152
211
Goodhue
Lex elem
FDC Elem

"""

##############################################################################
# Functions: Main() and Excel editing
##############################################################################


def main(folder_name):
    """ Runs the program to open browser and scrape """
    # Create variables to store site information
    total_cancelled = 0
    total_UA_sum = 0
    total_count = 0

    # Connect to functions class
    web = UnaccountedAttendanceScraping()

    # Open and enter dycd site
    web.generateBroswer()
    web.enter_DYCD(user, passw)

    # Navigate to third page and find UA report
    web.next_page()
    web.next_page()
    web.find_report('//*[@filename="UnaccountedForAttendance.rdl"]')

    # Set holding variable to save program area to save time
    web.global_prev(0)

    # Goal is to figure out which are summer workscopes or SY scopes
    sycompass_workscopes_build = []
    sumcompass_workscopes_build = []

    web.wait(2, 0)
    section = 2
    ws_list = web.fill_ua_forlist(start_date, section)

    # Ignore the <Select> in first element
    ws_list = ws_list[1:]

    # Look through the list and separate all with 9 in first number of date
    for i, ws in enumerate(ws_list, 1):
        if ws.split('-')[2][0] == str(9):
            sycompass_workscopes_build.append(i)
        else:
            sumcompass_workscopes_build.append(i)

    print(sycompass_workscopes_build)
    print(sumcompass_workscopes_build)

    # Create workscope list based on term
    if term == 'School Year':
        # Each workscope includes [program area, [workscopes]]
        compass_workscopes = [2, sycompass_workscopes_build]
        beacon_workscopes = [1, [1]]
        alit_workscopes = [3, [1, 2]]
        workscopes = [beacon_workscopes, alit_workscopes, compass_workscopes]
    else:
        # SUMMER
        compass_workscopes = [2, sumcompass_workscopes_build]
        beacon_workscopes = [1, [1]]
        workscopes = [beacon_workscopes, compass_workscopes]

    # Fill out report, download, and run excel function
    for section in workscopes:
        for workscope in section[1]:
            for i in range(3):  # three attempts before moving on to next
                try:
                    web.wait(2, 0)

                    # Fill and remember workscope name to put into Excel
                    name = web.fill_ua(start_date, section[0], workscope, sycompass_workscopes_build)

                    # Covert Excel to df
                    web.wait(4, 0)
                    df = get_UA_tab()
                    df = df.ffill()

                    # Create summary from df
                    cancelled, count_1, count_2 = web.create_summary(
                                                            start_date,
                                                            end_date,
                                                            name, section,
                                                            df,
                                                            workscope,
                                                            total_workscopes,
                                                            folder_name)
                    
                    # Add to total for YD_Summary
                    total_cancelled += cancelled
                    total_UA_sum += count_1
                    total_count += count_2
                except Exception as e:
                    print(e)
                    print(section, workscope, 'Not Working')
                    web.wait(3, 0)
                else:
                    break

    # Header for YD Summary
    YD_summary = [['Start Date:', str(start_date).split(' ')[0]],
                  ['End Date:', str(end_date).split(' ')[0]],
                  ['Report Run On:', str(date.today())],
                  ['', ''],
                  ['YD Summary:', ''],
                  ['Total Cancelled:', total_cancelled],
                  ['Total Unaccounted Attendance Sum*:', total_UA_sum],
                  ['Total Unaccounted Attendance Day & Activity Count**:', total_count]]

    YD_summary_df = pd.DataFrame(data=YD_summary, columns=['col1', 'col2'])

    # If exists, get UA sum from two weeks prior pull
    if first_round == False:
        old_folder_name = "A:/Office of Performance Management/Youth Division/Python Programs/DYCD Scraping/Backend/Unaccounted Attendance Output/" + str(start_date.date()) + ' to ' + str(end_date.date()) + '/' + 'UA_YDsummary_' + str(start_date).split(' ')[0] + '.xlsx'

        df = pd.read_excel(old_folder_name, skiprows=14)

        df['Revisited UA ' + str(start_date).split(' ')[0]] = df['Unaccounted Attendance Sum']
        revisited_ua = df['Revisited UA ' + str(start_date).split(' ')[0]]

    # Build dataframe with all school details
    total_df = pd.DataFrame(data=total_workscopes, columns=[
                                                    'Cohort',
                                                    'Site',
                                                    'Cancelled',
                                                    'Unaccounted Attendance Sum',
                                                    'Unnacounted Attendance Day & Activity Count'])
    total_df['# of Workscopes with Missing Attendance Data'] = total_df.apply(lambda row: 1 if row['Unaccounted Attendance Sum'] > 0 else 0, axis=1)
    total_df = total_df.sort_values(['Cohort', 'Site'])

    print(total_df)

    # Create cohort summary df
    cohorts_df = total_df['# of Workscopes with Missing Attendance Data'].gt(0).astype(int).groupby(total_df['Cohort']).sum()

    print(total_df)
    # No longer need this col, used above
    total_df = total_df.drop(['# of Workscopes with Missing Attendance Data'], axis=1)

    print(total_df)

    if first_round == True:
        # Create YD summary excel
        writer = pd.ExcelWriter(folder_name + 'UA_YDsummary_' + str(start_date).split(' ')[0] + '.xlsx')
    else: # subtract a week!
        # total_df = total_df.insert(4, 'Revisited UA ' + str(start_date).split(' ')[0], revisited_ua)
        writer = pd.ExcelWriter(folder_name + 'UA_YDsummary_Revisited_' + str(start_date).split(' ')[0] + '.xlsx')


    notes_df = pd.DataFrame(data=[
        '''* This is the total number of missing attendance entries for this time period and represents the number of data entries needed to be caught up. This does not include canceled days',\n
        ** This is the number of instances of missing attendance across days and activities/groups. This does not include canceled days.'''])

    # Write in each df to excel writer
    YD_summary_df.to_excel(writer, index=False, header=False)
    cohorts_df.to_excel(writer, startrow=len(YD_summary_df)+1)
    total_df.to_excel(writer, startrow=len(YD_summary_df)+len(cohorts_df)+3, index=False)
    notes_df.to_excel(writer, startrow=len(YD_summary_df)+len(cohorts_df)+len(total_df)+5, index=False, header=False)

    # Auto-adjust columns' width for total df
    for column in total_df:
        column_width = max(total_df[column].astype(str).map(len).max(), len(column))
        col_idx = total_df.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

    # Auto-adjust columns' width for yd sumary df col1
    column = 'col1'
    column_width = max(YD_summary_df[column].astype(str).map(len).max(), len(column))
    col_idx = YD_summary_df.columns.get_loc(column)
    writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

    writer.save()

    web.quit_browser()


def get_UA_tab():
    """ Left join eto app status to cohort df

    Returns
    -------
    df : dataframe
        df from Excel
    """

    # downloads_path = str(Path.home() / "Downloads")
    # file = str(downloads_path) + r'\Unaccounted Attendance.xlsx'
    # file = r'C:\Users\mrozanoff\Downloads\Unaccounted Attendance.xlsx'

    file = get_file()

    df = pd.read_excel(file, skiprows=21)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    os.remove(file)

    return df


##############################################################################
# GUi
##############################################################################

# to store list of sites information before making it into a df
total_workscopes = []
first_round = True

sg.theme('BlueMono')  # Add a touch of color

# All the stuff inside your window.
layout = [[sg.Push(),
            sg.CalendarButton('Select Start Date',  target='-IN1-', format='%m/%d/%Y'),
            sg.Input(key='-IN1-', size=(20, 1)), sg.Push(), sg.Text('Select Term:'),
            sg.Push(),
            sg.Combo(['School Year', 'Summer'], size=(10, 0), enable_events=True, key='-IN2-'),
            sg.Push()],
            [sg.Text('Enter DYCD Usename'), sg.InputText(key='-IN3-')],
            [sg.Text('Enter DYCD Password'), sg.InputText(password_char='*', key='-IN4-')],
            [sg.Button('Run Program'), sg.Push(), sg.Button('Cancel')],
            [sg.StatusBar("", size=(0, 1), key='-STATUS-')]]

# Create the Window
window = sg.Window('DYCD Unnacounted Attendance Scrape', layout)

# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    state = ''

    term = values['-IN2-']

    if event == sg.WIN_CLOSED or event == 'Cancel':
        break

    if event == 'Run Program':
        start_date = pd.to_datetime(values['-IN1-'])

        user = values['-IN3-']
        passw = values['-IN4-']

        if user and passw and term and start_date:
            state = "Login OK"
            end_date = start_date + timedelta(days=6)

            # Create date folder
            folder_name = "A:/Office of Performance Management/Youth Division/Python Programs/DYCD Scraping/Backend/Unaccounted Attendance Output/" + str(start_date.date()) + ' to ' + str(end_date.date()) + '/'

            # main(folder_name)

            # Run a second time, saving within the already created folder
            # so that can see sites that followed up on their UA
            first_round = False
            start_date = start_date - timedelta(weeks=2)
            end_date = start_date + timedelta(days=6)

            new_folder_name = folder_name + "Revisited_" + str(start_date.date()) + ' to ' + str(end_date.date()) + '/'
            folder_name = new_folder_name
            # print(folder_name)
            # quit()
            main(folder_name)

        else:
            state = "Login failed, please enter all inputs"

    window['-STATUS-'].update(state)

window.close()
