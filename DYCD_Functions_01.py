import time
import sys
import traceback
import chromedriver_autoinstaller
import pandas as pd
from datetime import datetime, timedelta, date
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC  

""" Selenium function needed to scrape through DYCD Connect

Currently supports pulls from:
-Unaccounted attendance
-The Attendance report
-ROP Beacon
-ROP CMS
-ROP CES
-ROP CYEP

TODO:
-Documentation
-Edit child class to work for weekly scrape
"""

class Webdriver:
    def generateBroswer(self):
        """ See if driver is installed and install it if not, get DYCD site """

        # Check if the current version of chromedriver exists and if it doesn't 
        # exist, download it automatically, then add chromedriver to path
        chromedriver_autoinstaller.install()

        options = webdriver.ChromeOptions()

        self.driver = webdriver.Chrome(options=options)
        self.driver.get("https://www.dycdconnect.nyc/Home/Login")
        self.driver.maximize_window()

    def quit_browser(self):
        """ Quit browser, needed? """
        self.driver.quit()

    def wait(self, t, explicit):
        """ Combine implicit and explicit waits. explicit may not work still.. """
        driver = self.driver
        time.sleep(t)
        WebDriverWait(driver, explicit).until(EC.invisibility_of_element((By.XPATH, '//*[@id="reportViewer_AsyncWait_Wait"]/table/tbody/tr/td[2]/div/a')), 'Timed out waiting for report to load')    

    def enter_DYCD(self, user, passw): 
        """ From the login page navigate to reports and open the report """
        driver = self.driver
        wait = self.wait

        # Input user and pass
        username = driver.find_element(By.ID, "UserName");
        password = driver.find_element(By.ID, "Password");

        username.send_keys(user)
        password.send_keys(passw)

        # Login
        buttons = driver.find_elements(By.TAG_NAME, 'button')
        buttons[1].click()
        wait(0.5, 0)

        # Click on PTS/EMS
        pts = WebDriverWait(driver, 1).until(EC.visibility_of_all_elements_located((By.CLASS_NAME, 'btn-group')))
        pts[1].click()

        # Switch to the new tab
        driver.switch_to.window(driver.window_handles[1])

        # Open menu and click reports
        menu = driver.find_element(By.CLASS_NAME, 'navBarTopLevelItem')
        button = menu.find_element(By.ID, 'TabMainMenu')
        button.click()
        wait(0.5, 0)

        nav_groups = driver.find_elements(By.CLASS_NAME, 'nav-subgroup')
        nav_groups[3].click()

        # Switch to iframe within reports page
        iframe = driver.find_element(By.XPATH, '//*[@id="contentIFrame0"]')
        driver.switch_to.frame(iframe)
        wait(1, 0)

    def next_page(self):
        """ Go to the next page in the reports viewer """
        driver = self.driver
        element = driver.find_element(By.XPATH, '//*[@id="_nextPageImg"]')
        driver.execute_script("arguments[0].click();", element)

    def find_report(self, report):
        """ Navigate the reports view and find given report
        
        Parameters
        ----------
        report : str
            XPATH for report
        """
        driver = self.driver
        wait = self.wait

        # Click tar (The Attendance Report)
        wait(0.5, 0)
        tar = driver.find_element(By.XPATH, report) 
        a = tar.find_element(By.TAG_NAME, 'a')

        actions = ActionChains(driver)
        actions.move_to_element(a).perform()
        wait(1, 0)

        driver.execute_script("arguments[0].click();", a)
        wait(1, 0)

        # Go to report builder page in new window
        driver.switch_to.window(driver.window_handles[2])
        driver.maximize_window()

        # Switch to iframe within report window
        driver.switch_to.frame(0)

    def select_element(self, xpath, n):
        """ Select element based on give XPATH
        
        Parameters
        ----------
        report : str
            XPATH for report
        """
        driver = self.driver
        select = Select(driver.find_element(By.XPATH, xpath))
        select.select_by_value(str(n))


    def retrive_workscope_list(self, xpath):
        driver = self.driver

        select = Select(driver.find_element(By.XPATH, xpath))
        
        compass_list = []
        for option in select.options:
            compass_list.append(option.text)


        return compass_list


    def fill_ua_forlist(self, start_date, program_area): #Run thru the Unaccounted attendance report
        driver = self.driver
        wait = self.wait
        select_element = self.select_element
        retrive_workscope_list = self.retrive_workscope_list


        name = 'na'

        wait(0, 3)
        select_element('//*[@id="reportViewer_ctl08_ctl04_ddValue"]', 2)
        wait(0, 3)

        if prev != program_area: #do this IF not the same as previous program.
            
            #program area
            wait(0, 3)
            select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', program_area)
            wait(0.5, 3)

            #provider
            select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', 1)
            wait(2, 3)

        #workscope
        compass_list = retrive_workscope_list('//*[@id="reportViewer_ctl08_ctl12_ddValue"]')
        wait(1, 3)

        return compass_list

    def select_workscope_element(self, xpath, n):
        driver = self.driver

        select = Select(driver.find_element(By.XPATH, xpath))

        select.select_by_value(str(n))

        return select.first_selected_option.text #, compass_list

    def download_report(self): #click view report
        driver = self.driver
        wait = self.wait

        wait(0, 3)
        driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl08_ctl00"]').click()

        wait(0, 200)
        driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl09"]/div/div[5]').click()

        wait(0, 200)
        driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl09_ctl04_ctl00_Menu"]/div[2]/a').click()  
     
    def global_prev(self, program_area): #speed things up in TAR by keeping certain elements plugged in
        global prev 
        prev = program_area


class UnaccountedAttendanceScraping(Webdriver):
    def fill_ua(self, start_date, program_area, workscope, workscope_list): #, ran_once): #Run thru the Unaccounted attendance report
        driver = self.driver
        wait = self.wait
        select_element = self.select_element
        select_workscope_element = self.select_workscope_element

        name = 'na'

        wait(0, 3)
        select_element('//*[@id="reportViewer_ctl08_ctl04_ddValue"]', 2)
        wait(0, 3)

        if prev != program_area: #do this IF not the same as previous program.
            
            #program area
            wait(0, 3)
            select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', program_area)
            wait(0.5, 3)

            #provider
            select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', 1)
            wait(2, 3)

        #workscope
        name= select_workscope_element('//*[@id="reportViewer_ctl08_ctl12_ddValue"]', workscope)
        wait(1, 3)

        #start date
        start = driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl08_ctl20_txtValue"]')
        start.clear()
        wait(0, 3)

        start.send_keys(str(start_date))
        start.send_keys(Keys.RETURN)
        wait(2, 3)

        #end date
        select_element('//*[@id="reportViewer_ctl08_ctl22_ddValue"]', 7)
        wait(2, 3)

        #group by activity or group. if elem, group. else activity
        # [6, 7, 8, 9, 10, 14, 15, 16, 20, 21, 22, 28, 29, 30, 31, 32, 35, 37, 39]
        # elem_scopes = [2, 4, 7, 9, 11, 13, 15, 17, 19, 23] #All except milbank elem
        # summer_groups = []
        # workscope_list = [6, 7, 8, 9, 10, 14, 15, 16, 20, 21, 22, 28, 29, 30, 31, 32, 35, 37, 39]
        # everything should select 1 (activity) for now. then filter those that need groups
        # CPE II [6] needs groups
        group_list = [8, 9, 10, 11, 16, 18, 22, 23, 24]
        """
        Groups:
        elem dyckman valley - 16
        dunlevy milbank elem - 20
        ps 5 - 8
        ps211 = 9
        cpe2 - 6
        fairmont - 7
        elem lex -22
        elem fred doug - 21


        """

        if workscope in group_list and program_area == 2:
            select_element('//*[@id="reportViewer_ctl08_ctl24_ddValue"]', 2)
        else:
            select_element('//*[@id="reportViewer_ctl08_ctl24_ddValue"]', 1)

        # Everything seems to be activity now
        # select_element('//*[@id="reportViewer_ctl08_ctl24_ddValue"]', 1)

        #Download report
        wait(2, 3)
        self.download_report()

        self.global_prev(program_area)
        wait(3, 0)

        return name


    def sort_cohort(self, name):
        bronx = [' Fairmont Neighborhood School', ' P.S. 211', ' I.S. X318 Math, Science & Technology Through Arts', ' P.S. 061 Francisco Oller', ' IS 219 New Venture School']
        manhattan = [' M.S. 319 - Maria Teresa', ' M.S. 324 - Patria Mirabal', ' P.S. 005 Ellen Lurie', ' P.S. 152 Dyckman Valley', ' Central Park East II', ' City College Academy of the Arts', ' Middle School 322', ' P.S. 008 Luis Belliard']
        centers = [' Frederick Douglass Center', ' Dunlevy Milbank Center', ' The Lexington Academy',  ' The Dunlevy Milbank Center', ' I.S. 061 William A Morris', ' The Childrens Aid Society']
        hs = [' Curtis High School']

        match name:
            case name if name in bronx:
                return 'Bronx'
            case name if name in manhattan:
                return 'Manhattan'
            case name if name in centers:
                return 'Centers'
            case name if name in hs:
                return 'High Schools'
            case _:
                return 'Literacy Services'


    def create_summary(self, start_date, end_date, name, section, df, workscope, total_workscopes, folder_name):
        #create date folder
        # folder_name = "A:/Office of Performance Management/Youth Division/Python Programs/DYCD Scraping/Backend/Unaccounted Attendance Output/" + str(start_date.date()) + ' to ' + str(end_date.date()) + '/'

        Path(folder_name).mkdir(parents=True, exist_ok=True)

        #calculate needed numbers for ua
        cancelled = len(df[df['Appointment Status'] == 'Canceled'])

        temp_df = df.loc[df['Appointment Status'] == 'Scheduled']
        count_1 = temp_df['Unaccounted Attendance'].sum()
        count_2 = temp_df['Unaccounted Attendance'][temp_df['Unaccounted Attendance'] > 0].count()

        if section[0] == 1:
            new_name = str(name.split('***')[-1])
        elif section[0] == 3:
            new_name = ' ' + ' '.join(name.split('-')[0:2]) + str(workscope)
        else:
            new_name = str(name.split('***')[-1]) + ' - ' + str(name.split()[1].split('-')[0])
            # temp_name = new_name.split(' - ')[0] + #create a fix for summer

        total_workscopes.append([self.sort_cohort(str(name.split('***')[-1])), new_name, cancelled, count_1, count_2])

        #summary df
        summary = [['Start Date:', str(start_date).split(' ')[0]],
                      ['End Date:', str(end_date).split(' ')[0]],
                      ['Report Run On:', str(date.today())],
                      ['',''],
                      ['Cancelled Activities / Groups:', cancelled],
                      ['Unaccounted Attendance Sum*:', count_1],
                      ['Unaccounted Attendance Day & Activity Count**:', count_2]
        ]

        summary_df = pd.DataFrame(data=summary)

        #start reading into excel
        writer = pd.ExcelWriter(folder_name + new_name.replace(' ', '_')[1:] + '_' + str(start_date).split(' ')[0] + '.xlsx') 

        notes_df = pd.DataFrame(data=['* This is the total number of missing attendance entries for this time period and represents the number of data entries needed to be caught up. This does not include canceled days',
        '** This is the number of instances of missing attendance across days and activities/groups. This does not include canceled days.'])

        #enter the dfs
        summary_df.to_excel(writer, index=False, header=False)
        df.to_excel(writer, startrow=len(summary_df)+1, index=False)
        notes_df.to_excel(writer, startrow=len(summary_df)+len(df)+3, index=False, header=False)

        # Auto-adjust columns' width
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

        writer.save()

        return cancelled, count_1, count_2

    
class WeeklyDashboardScrape(Webdriver):
    def global_prev_tar(self, program_area, program_type): #speed things up in TAR by keeping certain elements plugged in
        global prev 
        prev = [program_area, program_type]

    def separate(self, data):
        wait = self.wait

        data.append(['','']) #separate so final excel is cleaner
        wait(1, 0)
        self.switchto_report()
        wait(1, 0)

    def switchto_report(self): #goes from report editor window back to main dycd reports page
        driver = self.driver 

        driver.close() #no longer need report window, close it

        driver.switch_to.window(driver.window_handles[1]) # switch to reports tab

        #switch to iframe within reports page
        iframe = driver.find_element(By.XPATH, '//*[@id="contentIFrame0"]')
        driver.switch_to.frame(iframe)

    def fill_dates_TAR(self, start_date, end_date): #fill in start date
        wait = self.wait
        driver = self.driver 

        start = driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl08_ctl18_txtValue"]')
        start.clear()
        wait(1, 3)

        start.send_keys(str(start_date))
        start.send_keys(Keys.RETURN)
        wait(3, 3)

        end = driver.find_element(By.XPATH, '//*[@id="reportViewer_ctl08_ctl20_txtValue"]')
        end.clear()
        wait(2, 3)

        end.send_keys(str(end_date))
        end.send_keys(Keys.RETURN)

    def fill_in_TAR(self, program_area, program_type, workscope, start_date, end_date): #fill in report parameters:
        wait = self.wait
        select_element = self.select_element
        select_workscope_element = self.select_workscope_element

        if prev[0] != program_area or prev[1] != program_type: #do this IF not the same as previous program.
            # select_program_area(program_area)
            select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', program_area)
            wait(1, 3)

            # select_program_type(program_type)
            select_element('//*[@id="reportViewer_ctl08_ctl08_ddValue"]', program_type)
            wait(0.5, 3)

            # select_provider(1)
            select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', '1')
            wait(0.5, 3)

        # name = select_workscope(workscope)
        name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl12_ddValue"]', workscope)
        wait(2, 3)
        
        self.fill_dates_TAR(start_date, end_date)
        wait(6, 3)

        self.download_report()
        wait(3, 3)

        self.global_prev_tar(program_area, program_type)

        return name

    def fill_ROP_CES(self, workscope):
        wait = self.wait
        select_element = self.select_element
        select_workscope_element = self.select_workscope_element

        select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', '1')
        wait(0.5, 3)

        name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl08_ddValue"]', workscope)
        wait(0.5, 3)

        select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', '1')
        wait(3, 0)

        self.download_report()
        wait(3, 0)

        return name

    def fill_ROP_CMS(self, workscope):
        wait = self.wait
        select_element = self.select_element
        select_workscope_element = self.select_workscope_element

        select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', '1') #select provider name
        wait(0.5, 3)

        name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl08_ddValue"]', workscope) #select workscope and record name
        wait(0.5, 3)

        select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', '1') # period
        wait(3, 0)

        self.download_report()
        wait(3, 0)

        return name

    def fill_ROP_CYEP(self, program_type):
        wait = self.wait
        select_element = self.select_element
        select_workscope_element = self.select_workscope_element

        select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', program_type) # select program type
        wait(0.5, 3)

        select_element('//*[@id="reportViewer_ctl08_ctl08_ddValue"]', '1') # select fiscal year
        wait(0.5, 3)

        select_element('//*[@id="reportViewer_ctl08_ctl10_ddValue"]', '1') # select provider
        wait(0.5, 3)

        name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl12_ddValue"]', '1') #select workscope and record name
        wait(1, 0)

        self.download_report()
        wait(3, 0)

        return name

    def fill_ROP_B(self):
        wait = self.wait
        select_element = self.select_element
        select_workscope_element = self.select_workscope_element

        select_element('//*[@id="reportViewer_ctl08_ctl06_ddValue"]', '1') #select provider name
        wait(0.5, 3)

        name = select_workscope_element('//*[@id="reportViewer_ctl08_ctl08_ddValue"]', '1') #select workscope and record name
        wait(0.5, 3)

        self.download_report()
        wait(3, 0)

        return name
