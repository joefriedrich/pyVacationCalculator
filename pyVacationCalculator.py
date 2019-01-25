'''
	Title:  SharePoint Time Off Calculator 0.6
	Author:  Joe Friedrich
	License:  MIT
'''

import requests
import re
import xml.etree.ElementTree as ET
from requests_ntlm import HttpNtlmAuth
from datetime import datetime

def get_root_xml():
    '''
        Use requests to grab .xml file containing calendar data
        returns root of xml (fromstring or the file version[test version])
    '''
    raw_website = requests.get(
        "Sharepoint Website",
        auth=HttpNtlmAuth('Domain\\Username','Password'))
    raw_text = raw_website.text
    return ET.fromstring(raw_text)

def generate_item_from_root_xml(root):
	'''
		Write!
	'''
    tags = ['{http://schemas.microsoft.com/ado/2007/08/dataservices}Title',
        '{http://schemas.microsoft.com/ado/2007/08/dataservices}EventDate',
        '{http://schemas.microsoft.com/ado/2007/08/dataservices}EndDate']
    for element in root.iter():
        if element.tag in tags:
            yield element.text

def calculate_vacation_days(start_date, end_date):
    '''
        Takes start and end date of entry.
        --Verifies that the entry doesn't span multiple years.
        Returns number of days.
    '''
    if start_date.year != end_date.year:
        print("=========ERROR!  VACATION SPANS MULTIPLE YEARS======"
            "\nSTART:  " + start_date + " END:  " + end_date)
    else:
        total_days = end_date - start_date
        return total_days.days + 1
        
def build_data_from_xml(xml):
	'''
		Write!
	'''
    data = []
    names = []
    years = []
    for item in xml:
        vacation_item = re.split(r' - ', item.upper())
        if vacation_item[0] not in names:
            names.append(vacation_item[0])
        
        start_date_raw = next(xml_generator)
        start_date = datetime.strptime(start_date_raw[:10], '%Y-%m-%d')
        vacation_item.append(start_date)
                
        end_date_raw = next(xml_generator)
        end_date = datetime.strptime(end_date_raw[:10], '%Y-%m-%d')
        vacation_item.append(end_date)
        if end_date.year not in years:
            years.append(end_date.year)
            
        vacation_item.append(calculate_vacation_days(start_date, end_date))
        
        data.append(vacation_item)
    return data, names, years

def get_employee_from_user(names, years):
    '''
        Get employee name and employee year from user.
        (Change:  should verify year in Years and name in Names)
    '''
    while True:
        employee_name = input("\nPlease enter the employee's name:  ").upper()
        if employee_name in names:
            employee_year = int(input("Please enter the year to calculate:  "))
            if employee_year in years:
                return employee_name, employee_year	
        print('Employee or year not in list, please try again.')

def get_employee_time_data(vacation_data, employee_name, employee_year):
    '''
        Takes list of entries(vacation_data), name(user input), year(user input).
        Generator that returns the next item that matches employee name and year
    '''
    employee_time_data = []
    for item in vacation_data:
        if item[0] == employee_name:
            if item[2].year == employee_year:
                yield item
    
def output(time_type, time_data, employee_name):
    '''
        Write!
    '''
    print('--------' + time_type + ' TIME-------')
    total_time = 0
    for event in time_data:
        total_time = total_time + event[4]
        print(event[0] + ' was ' + event[1] + ' from ' + event[2].isoformat()[:10] +
        ' to ' + event[3].isoformat()[:10] + ' for ' + str(event[4]) + ' day(s).')
    print('Total ' + time_type + ' time for ' + employee_name +
            ' is ' + str(total_time))

#========================Begin Main Program============================
print('Retrieving, parsing, and organizing data.')
root = get_root_xml()
xml_generator = generate_item_from_root_xml(root)
vacation_data, names, years = build_data_from_xml(xml_generator)
stop = False

while(stop == False):
    employee_name, employee_year = get_employee_from_user(names, years)
    employee_calendar_data = list(get_employee_time_data(vacation_data, employee_name, employee_year))
    
    #replace these 2 snippits with a function
    employee_sick_data = []
    employee_vacation_data = []
    employee_wfh_data = [] #wfh = work from home

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

    output('SICK', employee_sick_data, employee_name)
    output('VACATION', employee_vacation_data, employee_name)
    output('Work From Home', employee_wfh_data, employee_name)
    
    pause = input('Press Enter to select another employee/year combination.')
