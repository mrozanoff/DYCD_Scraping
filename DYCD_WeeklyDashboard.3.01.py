# Import code that extracts information from downloaded Excel
import DYCD_ExcelFunctions as ga

# Import code that has the scraping functions
from DYCD_Functions_01 import WeeklyDashboardScrape

import time
import traceback

import pandas as pd
import PySimpleGUI as sg

from datetime import datetime
from dateutil.relativedelta import relativedelta

'''' This script collects the data for the DYCD Weekly Dashboard

It will open the DYCD website, login, and open the corresponsing reports.
These include the ROP reports, and the attendance reports. Then extracts the
data and compiles it into a single Excel

TODO:
-Finish Doc and flake8
-fix three eweks, needs to be three weeks back from previous monday, not on
whatever date im doing --?
'''

##############################################################################
### Main
##############################################################################


def main():
    try:
        # Connect to functions class which does most of the heavy lifting
        web = WeeklyDashboardScrape()

        # Open and enter dycd site
        web.generateBroswer()
        web.enter_DYCD(user, passw)
        web.next_page()

        # Begin the list that will be used to create the new Excel
        data.append(['Start Date: ', start_date])
        data.append(['End Date: ', end_date])
        data.append(['', ''])

        # Look at different reports depending if SY or Summer
        if term == 'School Year':
            # Rate of Participation BEACON
            web.find_report('//*[@filename="BEACON_ROP.rdl"]')
            data.append(['ROP Beacon:', ''])

            # Will try three times to pull the report
            for i in range(3):
                try:
                    name = web.fill_ROP_B()
                except Exception as e:
                    print(e)
                    time.sleep(3)
                else:
                    break

            time.sleep(3)

            try:  # If value does not exist (DYCD hasn't recorded yet), then just put n/a
                data.append([name + ' Middle School', ga.get_ROP_B('Sheet2', start_date)])
            except Exception as e:
                print(e)
                data.append([name, 'n/a'])

            try:
                data.append([name + ' High School', ga.get_ROP_B('Sheet3', start_date)])
            except Exception as e:
                print(e)
                data.append([name, 'n/a'])

            print(data)

            # Creates a space in Excel and switches to another report
            web.separate(data)

            web.find_report('//*[@filename="COMPASS_ROP_OVY.rdl"]')
            data.append(['ROP CYEP:', ''])

            # Rate of Participation CYEP
            for x in ROP_CYEP:
                for i in range(3):
                    try:
                        name = web.fill_ROP_CYEP(x)
                    except Exception as e:
                        print(e)
                        time.sleep(2)
                    else:
                        break

                time.sleep(2)
                try:  # Extract from Excel and store in data
                    if x == 1:
                        data.append([name, ga.get_ROP_CYEPe(start_date)])
                    else:
                        data.append([name, ga.get_ROP_CYEPhs(start_date)])
                except Exception as e:
                    print(e)
                    data.append([name, 'n/a'])

                print(data)

            web.separate(data)

            web.find_report('//*[@filename="COMPASS_ROP_Middle_Unstruct.rdl"]')
            data.append(['ROP CMS:', ''])

            # Rate of Participation CMS
            for x in ROP_CMS:
                for i in range(3):
                    try:
                        name = web.fill_ROP_CMS(x)
                    except Exception as e:
                        print(e)
                        print(traceback.format_exc())
                        time.sleep(3)
                    else:
                        break

                time.sleep(2)
                try:  # Extract from Excel and store in data
                    data.append([name, ga.get_ROP_CMS(start_date)])
                except Exception as e:
                    print(e)
                    data.append([name, 'n/a'])
                print(data)

            web.separate(data)

            web.find_report('//*[@filename="CompassElementary_ROP_New.rdl"]')
            data.append(['ROP CES:', ''])

            # Rate of Participation CES
            for x in ROP_CES:
                for i in range(3):
                    try:
                        name = web.fill_ROP_CES(x)
                    except Exception as e:
                        print(e)
                        print(traceback.format_exc())
                        time.sleep(3)
                    else:
                        break

                time.sleep(2)
                try:  # Extract from Excel and store in data
                    data.append([name, ga.get_ROP_CES(start_date)])
                except Exception as e:
                    print(e)
                    data.append([name, 'n/a'])

                print(data)

            web.separate(data)

        web.next_page()

        web.find_report('//*[@filename="TheAttendanceProc_rpt.rdl"]')
        data.append(['Attendance:', ''])

        # ##################################################################3
        #  # Goal is to figure out which are summer workscopes or work scopes
        # sycompass_workscopes_build = []
        # sumcompass_workscopes_build = []

        # web.wait(2, 0)
        # program_area =
        # ws_list = web.fill_in_TAR(program_area, program_type, workscope, start_date, end_date)

        # # Ignore the <Select> in first element
        # ws_list = ws_list[1:]

        # # Look through the list and separate all with 9 in first number of date
        # for i, ws in enumerate(ws_list, 1):
        #     if ws.split('-')[2][0] == str(9):
        #         sycompass_workscopes_build.append(i)
        #     else:
        #         sumcompass_workscopes_build.append(i)

        # print(sycompass_workscopes_build)
        # print(sumcompass_workscopes_build)
        # ####################################################################

        web.global_prev_tar(0, 0)  # Set prev to default nothing

        # Attendance
        for workscope in TAR_workscopes:
            for x in workscope[2]:  # x being program area
                time.sleep(1)
                for i in range(3):
                    try:  # Fill in the report
                        name = web.fill_in_TAR(workscope[0],
                                               workscope[1],
                                               x,
                                               start_date,
                                               end_date)
                    except Exception as e:
                        print(e)
                        time.sleep(3)
                    else:
                        break

                time.sleep(2)

                # Based on program area  use different extraction
                if workscope[0] == 3:
                    current = ga.get_attendance_COMPASS()
                    data.append([name, current])
                elif workscope[0] == 2:
                    beacon_tabs = ['Sheet2', 'Sheet3', 'Sheet4']

                    for tab in beacon_tabs:
                        try:
                            current = ga.get_attendance_Beacon(tab)
                            data.append([name + tab, current])
                        except Exception as e:
                            print(e)
                            data.append([name + tab, 'Sheet not Found'])
                            # print(traceback.format_exc())
                            time.sleep(3)
                else:
                    current = ga.get_attendance_aLit()
                    data.append([name, current])

                print(data)

        data.append(['Cumulative Attendance:', ''])

        web.global_prev_tar(0, 0)  # set prev to default nothing

        # Cumulative Attendance, identical to above apart from dates
        for workscope in TAR_workscopes:
            for x in workscope[2]:
                for i in range(3):
                    try:
                        name = web.fill_in_TAR(workscope[0],
                                               workscope[1],
                                               x,
                                               school_start_date,
                                               todays_date)
                    except Exception as e:
                        print(e)
                        # print(traceback.format_exc())
                        time.sleep(3)
                    else:
                        break

                time.sleep(2)

                try:
                    if workscope[0] == 3:
                        current = ga.get_attendance_COMPASS()
                        data.append([name, current])
                    elif workscope[0] == 2:
                        beacon_tabs = ['Sheet2', 'Sheet3', 'Sheet4']
                        for tab in beacon_tabs:
                            current = ga.get_attendance_Beacon(tab)
                            data.append([name + tab, current])
                    else:
                        current = ga.get_attendance_aLit()
                        data.append([name, current])
                except Exception as e:
                    print(e)

                print(data)

    except Exception as e:
        print(e)

    # Create df from list
    df = pd.DataFrame(data=data)

    save_file = 'A:/Office of Performance Management/Youth Division/Python Programs/DYCD Scraping/Backend/Weekly Dashboard Output/DYCDattendance_list_' + timestr + '.xlsx'
    df.to_excel(save_file, index=False, header=False)

    # Goodbye and thanks for all the fish
    web.quit_browser()


##############################################################################
### GUI
##############################################################################

sg.theme('BlueMono')  # Add a touch of color
# All the stuff inside your window.
layout = [[sg.Text('Enter Term'), sg.Combo(['School Year', 'Summer'] , key='-IN0-'), sg.Push()],
          [sg.Push(), sg.CalendarButton('Select Start Date', target='-IN1-', format='%m/%d/%Y'), sg.Input(key='-IN1-', size=(20, 1)), sg.Push()],
          [sg.Push(), sg.CalendarButton('Select End Date', target='-IN2-', format='%m/%d/%Y'), sg.Input(key='-IN2-', size=(20, 1)), sg.Push()],
          [sg.Text('Enter DYCD Usename'), sg.InputText(key='-IN3-')],
          [sg.Text('Enter DYCD Password'), sg.InputText(password_char='*', key='-IN4-')],
          # [sg.Checkbox("Compass", default=True, key='-CB1-', enable_events=True),
          #     sg.Checkbox("Beacon", default=True, key='-CB2-', enable_events=True),
          #     sg.Checkbox("Alit", default=True, key='-CB3-', enable_events=True)],
          [sg.Button('Run Program'), sg.Push(), sg.Button('Cancel')]]

# Create the Window
window = sg.Window('DYCD Weekly Dashboard Scrape', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    term = values['-IN0-']
    start_date = values['-IN1-']
    end_date = values['-IN2-']
    user = values['-IN3-']
    passw = values['-IN4-']

    # Input start date
    todays_date = datetime.now()

    timestr = time.strftime("%Y-%m-%d_%H-%M-%S")

    if term == 'School Year':

        sep_1 = pd.to_datetime("9/1/" + str(todays_date.year))

        if todays_date > sep_1:
            school_start_date = sep_1
        else:
            school_start_date = sep_1 - relativedelta(years=1)

        # TAR workscopes: [program area, program type, workscopes]
        TARc_elementary_workscopes = [3, 1, [6, 7, 8, 9, 10, 14, 15, 16, 20, 21, 22]]
        TARc_explore_workscopes = [3, 2, [1]]
        TARc_high_workscopes = [3, 3, [1]]
        TARc_middle_workscopes = [3, 5, [6, 7, 8, 9, 10, 12, 14, 16]]
        TARb_Beacon = [2, 1, [1]]
        TARl_aLit = [8, 1, [1, 2]]
        TARl_AdultLit_AHdis = [7, 3, [1]]
        TARl_AdultLit_BEdis = [7, 5, [1]]

        TAR_workscopes = [TARb_Beacon,
                          TARl_aLit,
                          TARc_elementary_workscopes,
                          TARc_middle_workscopes,
                          TARc_high_workscopes,
                          TARc_explore_workscopes]
        # , TARl_aLit, TARl_AdultLit_AHdis, TARl_AdultLit_BEdis] also removed becaone, explore, high school

        ROP_CES = [6, 7, 8, 9, 10, 14, 15, 16, 20, 21, 22]
        ROP_CMS = [6, 7, 8, 9, 10, 12, 14, 16]
        ROP_CYEP = [1, 2]

    elif term == 'Summer':  # Summer Version!

        school_start_date = pd.to_datetime("7/1/" + str(todays_date.year))

        # TAR workscopes: [program area, program type, workscopes]
        TARc_elementary_workscopes = [3, 1, [1, 2, 3, 4, 5, 11, 12, 13, 17, 18, 19]]
        TARc_middle_workscopes = [3, 5, [6, 7, 8, 9, 10, 12, 14, 16]]
        TARb_Beacon = [2, 1, [1]]

        TAR_workscopes_summer = [TARb_Beacon, TARc_elementary_workscopes, TARc_middle_workscopes]  # TARl_AdultLit_AHdis, TARl_AdultLit_BEdis]

        TAR_workscopes = TAR_workscopes_summer

    if event == 'Run Program':
        data = []  # Keep a list to put into final df
        main()

    if event == sg.WIN_CLOSED or event == 'Cancel':
        break

window.close()
