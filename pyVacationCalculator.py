'''
	Title:  SharePoint Time Off Calculator 0.5
	Author:  Joe Friedrich
	License:  MIT
'''

import openpyxl
import os
import re
from datetime import date
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def get_excel_data():
	'''
		Write!
	'''
	print('Please hit the enter key when you are ready to')
	pause = input('select the sharepoint calendar file.')
	window = Tk()
	window.withdraw()
	file_name = askopenfilename()

	excel_data = openpyxl.load_workbook(file_name, read_only = True)
	return excel_data

def get_employee_from_user():
	'''
		Write!
	'''
    employee_name = input("\nPlease enter the employee's name:  ").upper()
    employee_year = int(input("Please enter the year to calculate:  "))
    return employee_name, employee_year

def is_it_time_to_stop():
	'''
		Write!
	'''
	stop = False
	stopping = input("\nShall we end the program?  ")
	if stopping in ('y', 'Y', 'yes', 'Yes','YEs','YES'):
		stop = True
	return stop

def calculate_vacation_days(start_date, end_date):
	'''
		Write!
	'''
    if start_date.year != end_date.year:
        print("=========ERROR!  VACATION SPANS MULTIPLE YEARS======"
             "\nSTART:  " + start_date + " END:  " + end_date)
    else:
        total_days = end_date - start_date
        return total_days.days + 1

def get_employee_time_data(vacation_data, employee_name, employee_year):
	'''
		Write!
	'''
    employee_time_data = []
    for item in vacation_data:
        if item[0] == employee_name:
            if item[2].year == employee_year:
                yield item

#for name in names, create a list of sets.
#the sets contain (date type, dates)

def parse_raw_data(excel_data):
	'''
		Write!
	'''
	names = []  #empty list to hold unique names
	vacation_data = [] #empty list to hold vacation data
	for row in excel_data['owssvr'].rows:
		if row[0].value != 'Title':
			vacation_item = re.split(r' - ', row[0].value.upper())    #split on this when looking at names field, front is name.  back is type.
			if vacation_item[0] not in names:
				names.append(vacation_item[0])
				
			date_format = re.compile(r'\d+')
			find_date = date_format.findall(str(row[1].value))
			start_date = date(int(find_date[0]), int(find_date[1]), int(find_date[2]))
			find_date = date_format.findall(str(row[2].value))
			end_date = date(int(find_date[0]), int(find_date[1]), int(find_date[2]))
			
			vacation_days = calculate_vacation_days(start_date, end_date) #this will provide a number that we will append to calulate vacation days.
			#breaks on year seperation and asks you to fix it.
			
			vacation_item.append(start_date) #append start date to the list
			vacation_item.append(end_date)
			vacation_item.append(vacation_days)
			vacation_data.append(vacation_item)
			
	return vacation_data
	
def output(time_type, time_data, employee_name):
	'''
		Write!
	'''
	print('--------' + time_type + ' TIME-------')
	total_time = 0
	for event in time_data:
		total_time = total_time + event[4]
		print(event[0] + ' was ' + event[1] + ' from ' + event[2].isoformat() +
		' to ' + event[3].isoformat() + ' for ' + str(event[4]) + ' day(s).')
	print('Total ' + time_type + ' time for ' + employee_name +
			' is ' + str(total_time))

#========================Begin Main Program============================

raw_calendar_data = get_excel_data()
calendar_data = parse_raw_data(raw_calendar_data)
stop = False

while(stop == False):

	employee_name, employee_year = get_employee_from_user()
	employee_calendar_data = list(get_employee_time_data(calendar_data, employee_name, employee_year))
	
	#replace these 2 snippits with a function
	employee_sick_data = []
	employee_vacation_data = []
	employee_wfh_data = [] #wfh = work from home

	#should be function  (takes employee_calendar_data, and list of key words?)
	#start with a value of 1
	#rewrite to change the value if the first char is a number
	#--remove first 3 values (throw away last [space])
	#use the rest to determine type of data
	for item in employee_calendar_data:
		if item[1][0].isdigit():
			item[4] = int(item[1][0]) / int(item[1][2]) #replace 1 with decimal
			item[1] = item[1][4:] #removes number and space
		if item[1] in ('SICK', 'OUT'):
			employee_sick_data.append(item)
		elif item[1] in ('VACATION', 'DAY'):
			employee_vacation_data.append(item)
		elif item[1] == 'WFH':
			employee_wfh_data.append(item)

	#gather these into a function
	output('SICK', employee_sick_data, employee_name)
	output('VACATION', employee_vacation_data, employee_name)
	output('Work From Home', employee_wfh_data, employee_name)

	stop = is_it_time_to_stop()