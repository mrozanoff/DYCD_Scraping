import win32com.client
import pandas as pd
from datetime import datetime, timedelta, date
import PySimpleGUI as sg  

"""

TODO:
Make interface allow selection of which workscopes to send to
"""

#--------------------------------------------Email Data

year = 24

summer_directory = {  # Too different to code, easier to rewrite
'Central_Park_East_II_-_FY' + str(24) + '_Summer_-_Elementary_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'jeanettet@childrensaidnyc.org',
	'Data Coordinator' : 'vnunez@childrensaidnyc.org'
	},
'Central_Park_East_II_-_FY' + str(24) + '_Summer_-_Middle_' : { ##############################middle
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'jeanettet@childrensaidnyc.org',
	'Data Coordinator' : 'vnunez@childrensaidnyc.org'
	},
'Dunlevy_Milbank_Center_-_Elementary_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'casperl@childrensaidnyc.org',
	'Data Coordinator' : 'lashandaf@childrensaidnyc.org'
	},
'Fairmont_Neighborhood_School_-_FY' + str(24) + '_Summer_-_Elementary_' : {
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'garyp@childrensaidnyc.org',
	'Data Coordinator' : 'nrivera@childrensaidnyc.org',
	'Program Director' : 'lcandelario@childrensaidnyc.org'
	},
'Frederick_Douglass_Center_-_Elementary_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'amyh@childrensaidnyc.org',
	'Data Coordinator' : 'tjenkins2@childrensaidnyc.org'
	},
'Frederick_Douglass_Center_-_Middle_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'amyh@childrensaidnyc.org',
	'Data Coordinator' : 'tjenkins2@childrensaidnyc.org'
	},
'I.S._061_William_A_Morris_-_FY' + str(24) + '_Summer_-_Middle_' : { ####################IS 61
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'ilenep@childrensaidnyc.org',
	'Data Coordinator' : 'katiel@childrensaidnyc.org'
	},
'IS_X318_Math,_Science_&_Technology_Through_Arts_(12X318)_-_FY' + str(24) + '_Summer_-_Middle_' : {
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'sandra.romero@childrensaidnyc.org',
	'Data Coordinator' : 'kpagan@childrensaidnyc.org',
	'Data Coordinator2' : 'jsalcedo@childrensaidnyc.org',
	'Youth Advocate' : 'dcampos@childrensaidnyc.org',
	},
'IS_219_New_Venture_School_' : {
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'jhernandez@childrensaidnyc.org',
	'Data Coordinator' : 'jennifera@childrensaidnyc.org',
	'Beacon Director' : 'tsimpson@childrensaidnyc.org',
	'Beacon Assistant Director' : 'nruiz@childrensaidnyc.org'
	},
'M.S._324_-_Patria_Mirabal_-_FY' + str(24) + '_Summer_-_Middle_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'tracizg@childrensaidnyc.org',
	'Afterschool Program Director1' : 'mminaya@childrensaidnyc.org',
	'Afterschool Program Director2' : 'valmonte@childrensaidnyc.org'
	},
'Middle_School_322_-_FY' + str(24) + '_Summer_-_Middle_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'migdaliac@childrensaidnyc.org',
	'Data Coordinator' : 'jdeorales@childrensaidnyc.org'
	}, 
'P.S._005_Ellen_Lurie_-_FY' + str(24) + '_Summer_-_Elementary_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'madelyng@childrensaidnyc.org',
	'Data Coordinator' : 'kbolanos@childrensaidnyc.org'
	},
'P.S._008_Luis_Belliard_-_FY' + str(24) + '_Summer_-_Elementary_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'arneryr@childrensaidnyc.org',
	'Data Coordinator' : ''
	},
'P.S._061_Francisco_Oller_-_FY' + str(24) + '_Summer_-_Elementary_' : {
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'mariap@childrensaidnyc.org',
	'Data Coordinator' : 'iruiz@childrensaidnyc.org'
	},
'P.S._152_Dyckman_Valley_-_FY' + str(24) + '_Summer_-_Elementary_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'gconcepcion@childrensaidnyc.org',
	'Data Coordinator' : 'ysanchez2@childrensaidnyc.org'
	},
'P.S._211_-_FY' + str(24) + '_Summer_-_Elementary_' : {
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'sandra.romero@childrensaidnyc.org',
	'Data Coordinator' : 'kpagan@childrensaidnyc.org',
	'Data Coordinator2' : 'jsalcedo@childrensaidnyc.org',
	'Youth Advocate' : 'dcampos@childrensaidnyc.org',
	},
'The_Childrens_Aid_Society_-_Elementary_' : { ###########Goodhue
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'ilenep@childrensaidnyc.org',
	'Data Coordinator' : 'katiel@childrensaidnyc.org'
	},
'The_Dunlevy_Milbank_Center_-_Middle_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'casperl@childrensaidnyc.org',
	'Data Coordinator' : 'lashandaf@childrensaidnyc.org'
	},
'The_Lexington_Academy_-_FY' + str(24) + '_Summer_-_Elementary_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'dgiordano@childrensaidnyc.org',
	'Data Coordinator' : 'jackelynr@childrensaidnyc.org'
	},
'The_Lexington_Academy_-_FY' + str(24) + '_Summer_-_Middle_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'dgiordano@childrensaidnyc.org',
	'Data Coordinator' : 'jackelynr@childrensaidnyc.org'
	}
}

directory = {
'766623C_Adolescent_Literacy1_' : { #WYJR
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'sandra.romero@childrensaidnyc.org',
	'Data Coordinator' : 'kpagan@childrensaidnyc.org'
	},
'766624C_Adolescent_Literacy2_' : { #SU
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'migdaliac@childrensaidnyc.org',
	'Data Coordinator' : 'jdeorales@childrensaidnyc.org'
	},
# '931892U_Adult_Literacy_3_' : {
# 	'Deputy Director' : 'myrnat@childrensaidnyc.org',
# 	'Site Director' : 'migdaliac@childrensaidnyc.org',
# 	'Data Coordinator' : 'jdeorales@childrensaidnyc.org'
# 	},
# '931892U_Adult_Literacy_4_' : {
# 	'Deputy Director' : 'myrnat@childrensaidnyc.org',
# 	'Site Director' : 'migdaliac@childrensaidnyc.org',
# 	'Data Coordinator' : 'jdeorales@childrensaidnyc.org'
# 	},
'Central_Park_East_II_-_Elementary_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'jeanettet@childrensaidnyc.org',
	'Data Coordinator' : 'vnunez@childrensaidnyc.org'
	},
'Central_Park_East_II_-_Middle_' : { ##############################middle
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'jeanettet@childrensaidnyc.org',
	'Data Coordinator' : 'vnunez@childrensaidnyc.org'
	},
'City_College_Academy_of_the_Arts_-_Middle_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'migdaliac@childrensaidnyc.org',
	'Data Coordinator' : 'jdeorales@childrensaidnyc.org'
	},
# 'Curtis_High_School_-_High_' : {
# 	'Deputy Director' : 'courtneyc@childrensaidnyc.org',
# 	'Site Director' : 'cmarks@childrensaidnyc.org',
# 	'Data Coordinator' : 'rtous@childrensaidnyc.org'
# 	},
'Dunlevy_Milbank_Center_-_Elementary_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'casperl@childrensaidnyc.org',
	'Data Coordinator' : 'lashandaf@childrensaidnyc.org'
	},
'Fairmont_Neighborhood_School_-_Elementary_' : {
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'garyp@childrensaidnyc.org',
	'Data Coordinator' : 'nrivera@childrensaidnyc.org',
	'Program Director' : 'lcandelario@childrensaidnyc.org'
	},
'Frederick_Douglass_Center_-_Elementary_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'amyh@childrensaidnyc.org',
	'Data Coordinator' : 'tjenkins2@childrensaidnyc.org'
	},
'Frederick_Douglass_Center_-_Middle_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'amyh@childrensaidnyc.org',
	'Data Coordinator' : 'tjenkins2@childrensaidnyc.org'
	},
'I.S._061_William_A_Morris_-_Middle_' : { ####################IS 61
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'ilenep@childrensaidnyc.org',
	'Data Coordinator' : 'katiel@childrensaidnyc.org'
	},
'I.S._X318_Math,_Science_&_Technology_Through_Arts_-_Middle_' : {
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'sandra.romero@childrensaidnyc.org',
	'Data Coordinator' : 'kpagan@childrensaidnyc.org',
	'Data Coordinator2' : 'jsalcedo@childrensaidnyc.org',
	'Youth Advocate' : 'dcampos@childrensaidnyc.org',
	},
'IS_219_New_Venture_School_' : {
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'jhernandez@childrensaidnyc.org',
	'Data Coordinator' : 'jennifera@childrensaidnyc.org',
	'Beacon Director' : 'tsimpson@childrensaidnyc.org',
	'Beacon Assistant Director' : 'nruiz@childrensaidnyc.org'
	},
'M.S._319_-_Maria_Teresa_-_Middle_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'tracizg@childrensaidnyc.org',
	'Afterschool Program Director1' : 'mminaya@childrensaidnyc.org',
	'Afterschool Program Director2' : 'valmonte@childrensaidnyc.org'
	},
'M.S._324_-_Patria_Mirabal_-_Middle_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'tracizg@childrensaidnyc.org',
	'Afterschool Program Director1' : 'mminaya@childrensaidnyc.org',
	'Afterschool Program Director2' : 'valmonte@childrensaidnyc.org'
	},
'Middle_School_322_-_Middle_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'migdaliac@childrensaidnyc.org',
	'Data Coordinator' : 'jdeorales@childrensaidnyc.org'
	}, 
'P.S._005_Ellen_Lurie_-_Elementary_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'madelyng@childrensaidnyc.org',
	'Data Coordinator' : 'kbolanos@childrensaidnyc.org'
	},
'P.S._008_Luis_Belliard_-_Elementary_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'arneryr@childrensaidnyc.org',
	'Data Coordinator' : ''
	},
'P.S._061_Francisco_Oller_-_Elementary_' : {
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'mariap@childrensaidnyc.org',
	'Data Coordinator' : 'iruiz@childrensaidnyc.org'
	},
'P.S._152_Dyckman_Valley_-_Elementary_' : {
	'Deputy Director' : 'myrnat@childrensaidnyc.org',
	'Site Director' : 'gconcepcion@childrensaidnyc.org',
	'Data Coordinator' : 'ysanchez2@childrensaidnyc.org'
	},
'P.S._211_-_Elementary_' : {
	'Deputy Director' : 'rcope@childrensaidnyc.org',
	'Site Director' : 'sandra.romero@childrensaidnyc.org',
	'Data Coordinator' : 'kpagan@childrensaidnyc.org',
	'Data Coordinator2' : 'jsalcedo@childrensaidnyc.org',
	'Youth Advocate' : 'dcampos@childrensaidnyc.org',
	},
'The_Childrens_Aid_Society_-_Elementary_' : { ###########Goodhue
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'ilenep@childrensaidnyc.org',
	'Data Coordinator' : 'katiel@childrensaidnyc.org'
	},
'The_Dunlevy_Milbank_Center_-_Middle_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'casperl@childrensaidnyc.org',
	'Data Coordinator' : 'lashandaf@childrensaidnyc.org',
	'Data Specialist': 'mbrown@childrensaidnyc.org'
	},
'The_Lexington_Academy_-_Elementary_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'dgiordano@childrensaidnyc.org',
	'Data Coordinator' : 'jackelynr@childrensaidnyc.org'
	},
# 'The_Lexington_Academy_-_Explore_' : {
# 	'Deputy Director' : 'vitoi@childrensaidnyc.org',
# 	'Site Director' : 'dgiordano@childrensaidnyc.org',
# 	'Data Coordinator' : 'cworko@childrensaidnyc.org'
# 	},
'The_Lexington_Academy_-_Middle_' : {
	'Deputy Director' : 'vitoi@childrensaidnyc.org',
	'Site Director' : 'dgiordano@childrensaidnyc.org',
	'Data Coordinator' : 'jackelynr@childrensaidnyc.org'
	}
}

#--------------------------------------------------

def main(_start_date, directory, _send_display, term):
	
	if term == 'Summer':
		directory = summer_directory


	def email_action():
		if _send_display == 'Display Emails First':
			newmail.Display()
		else:
			# pass
			newmail.Send()

	#Same dates as in UA
	start_date = _start_date

	end_date = start_date + timedelta(days=6)

	#location based on dates
	folder_name = "A:/Office of Performance Management/Youth Division/Python Programs/DYCD Scraping/Backend/Unaccounted Attendance Output/" + str(start_date) + ' to ' + str(end_date) + '/'

	ol = win32com.client.Dispatch("outlook.application")

	olmailitem = 0x0 #size of the new email

	# #sending out the site summaries
	for school in directory:
		print(school)
		newmail = ol.CreateItem(olmailitem)

		attachment = folder_name + str(school) + str(start_date) + '.xlsx'   #input attachement name
		emails = '; '.join(list(directory[school].values()))
		print(attachment)
		newmail.Subject = 'Unaccounted Attendance for week of ' + str(start_date) 
		print(attachment)
		newmail.Attachments.Add(attachment)
		newmail.To = emails
		newmail.Body = "Hello,\n\nI have updated the code for the school year, please point out any errors or changes that are needed. Thank you.\n\nThis is a weekly automated email summarizing the unaccounted attendance for your site. Please use this resource to get ahead of the DYCD lock.\n\nPlease contact part-time data entry assistance if you would like assistance catching up:\nCarlos Diaz: cdiaz@childrensaidnyc.org\nAnisa Bomani: abomani@childrensaidnyc.org\n\nIf I missed anybody in the email list / you would like somebody added please let me know.\n\nLet me know if you have any questions.\n\nThank you,\nMatthew" 

		email_action()
		
	# YD Overall summary sent to leadership and deputies
	newmail = ol.CreateItem(olmailitem)

	attachment = folder_name + 'UA_YDsummary_' + str(start_date) + '.xlsx'   #input attachement name _Revisited
	# revisit_attachment = attachment # figure thi out
	emails = 'sjonas@childrensaidnyc.org; myrnat@childrensaidnyc.org; rcope@childrensaidnyc.org; jesikac@childrensaidnyc.org; vitoi@childrensaidnyc.org; moniquew@childrensaidnyc.org; courtneyc@childrensaidnyc.org; jburke@childrensaidnyc.org'

	newmail.Subject = 'Unaccounted Attendance for week of ' + str(start_date) 
	newmail.Attachments.Add(attachment)
	newmail.To = emails
	newmail.Body = "Hello,\n\nThis is a weekly automated email summarizing the unaccounted attendance for the Youth Division. \n\nLet me know if you have any questions.\n\nThank you,\nMatthew" 

	email_action()


sg.theme('BlueMono') # Add a touch of color
# All the stuff inside your window.
layout = [  [sg.Text('Please enter the SAME date as you just did for the unaccounted attendance pull. If you pulled data for 5.22-5.28 enter 5.22.')],
			[sg.Push(), sg.CalendarButton('Select Start Date',  target='-IN1-', format='%m/%d/%Y'), sg.Input(key='-IN1-', size=(20,1)), sg.Push(), sg.Combo(['Display Emails First', 'Send Emails Automatically'], default_value='Display Emails First', key='-IN2-'), sg.Push(), sg.Combo(['School Year', 'Summer'], size=(10, 0), key='-IN3-'), sg.Push()],
            [sg.Button('Run Program'), sg.Push(), sg.Button('Cancel')] ]

# Create the Window
window = sg.Window('DYCD Weekly Dashboard Scrape', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    start_date = pd.to_datetime(values['-IN1-'])
    send_display = values['-IN2-']
    term = values['-IN3-']
    print(send_display)
    # Path(folder_name).mkdir(parents=True, exist_ok=True)

    if event == 'Run Program':
        main(start_date.date(), directory, send_display, term)

    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break

window.close()