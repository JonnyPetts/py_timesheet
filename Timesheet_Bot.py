#!/usr/bin/env python3
import os
import sys
import site
import threading
import time
from datetime import datetime
from datetime import timedelta

__author__ = "Jonny Petts"
__version__ = "1.1"
__license__ = "GNU GPLv3"

try:
    __import__('imp').find_module('openpyxl')
    print('')
    print('Success OpenPyXl Detected')
    # Make things with supposed existing module
except ImportError:
    pass
    print('')
    print('You do not have OpenpyXl installed - it will now install')
    print('')
    print(site.getsitepackages())
    time.sleep(2)
    
import openpyxl


#Pointless spin effect
class Spinner:
    busy = False
    delay = 0.1

    @staticmethod
    def spinning_cursor():
        while 1: 
            for cursor in '|/-\\': yield cursor

    def __init__(self, delay=None):
        self.spinner_generator = self.spinning_cursor()
        if delay and float(delay): self.delay = delay

    def spinner_task(self):
        while self.busy:
            sys.stdout.write(next(self.spinner_generator))
            sys.stdout.flush()
            time.sleep(self.delay)
            sys.stdout.write('\b')
            sys.stdout.flush()

    def start(self):
        self.busy = True
        threading.Thread(target=self.spinner_task).start()

    def stop(self):
        self.busy = False
        time.sleep(self.delay)


print("")
print("Welcome to Timesheet Bot 1.1")
print ('\033[91m' + "I now will ask you a few questions about your work week." + '\033[0m')
time.sleep(3)

# Location of spreadsheet

file = str(input('Please drag and drop the Timesheet Folder here : '))
path = file[:-1]
os.chdir(path)
wb = openpyxl.load_workbook('DC.xlsx')
FMT = '%H:%M'
sheet = wb['TIMESHEET']

# Name of the individual
print("")
name = str(input('What is your full name? '))
sheet['B9'] = name

# Company name
print("")
company = str(input('What is your company name? '))
sheet['J9'] = company

# The individuals job on set
print("")
job = str(input('What is your job title? '))
sheet['B14'] = job

# The dept that the individual is in
print("")
dept = str(input('What dept. are you in? '))
sheet['T14'] = dept

# The unit that they are in (Main or Second unit)
print("")
unit = str(input('Which Unit are you ? M, SU, UW, SP, VFX '))
sheet['B20'] = unit
sheet['B21'] = unit
sheet['B22'] = unit
sheet['B23'] = unit
sheet['B24'] = unit
sheet['G20'] = ('W')
sheet['G21'] = ('W')
sheet['G22'] = ('W')
sheet['G23'] = ('W')
sheet['G24'] = ('W')

# Are they a daily or weekly hire
print("")
hire = str(input('Are you Daily or Weekly hire? (d or w) ').lower())
if hire == ("d"):
	sheet['T9'] = ('X')
else :
	sheet['V9'] = ('X')	
	
	
# Which sites did they work at ?
print("")
same_site = str(input('Did you work at the location all week? y or n ').lower())
if same_site == ("y"):
	b1 = (input('What was you site location? '))
	sheet['N20'] = b1
	sheet['N21'] = b1
	sheet['N22'] = b1
	sheet['N23'] = b1
	sheet['N24'] = b1
else :
	msite = str(input('What Site did you work at on Monday?'))
	sheet['N20'] = msite
	tusite = str(input('What Site did you work at on Tuesday?'))
	sheet['N21'] = tusite
	wsite = str(input('What Site did you work at on Wednesday?'))
	sheet['N22'] = wsite
	thsite = str(input('What Site did you work at on Thursday?'))
	sheet['N23'] = thsite
	fsite = str(input('What Site did you work at on Friday?'))
	sheet['N24'] = fsite


# Calculate their working time
print("")
easy_week = str(input('Did you work the same times all week? y or n ').lower())
if easy_week == ("y"):
	s1 = (input('What was your start time ? >>> eg 08:00 '))
	sheet['H20'] = s1
	sheet['H21'] = s1
	sheet['H22'] = s1
	sheet['H23'] = s1
	sheet['H24'] = s1
	s2 = (input('What was your end time ? >>> eg 18:00 '))
	sheet['J20'] = s2
	sheet['J21'] = s2
	sheet['J22'] = s2
	sheet['J23'] = s2
	sheet['J24'] = s2
	adelta = datetime.strptime(s2, FMT) - datetime.strptime(s1, FMT)
	if adelta.days < 0:	
		adelta = timedelta(days=0,seconds=tdelta.seconds, microseconds=tdelta.microseconds)
	sheet['L20'] = adelta
	sheet['L21'] = adelta
	sheet['L22'] = adelta
	sheet['L23'] = adelta
	sheet['L24'] = adelta	
else :
	m1 = (input('What was your start time on Monday ? HH:MM '))
	m2 = (input('What was your end time on Monday ? HH:MM '))
	sheet['H20'] = m1
	sheet['J20'] = m2
	mdelta = datetime.strptime(m2, FMT) - datetime.strptime(m1, FMT)
	if mdelta.days < 0:	
		mdelta = timedelta(days=0,seconds=mdelta.seconds, microseconds=mdelta.microseconds)
	sheet['L20'] = mdelta
	
	tu1 = (input('What was your start time on Tuesday ? HH:MM '))
	tu2 = (input('What was your end time on Tuesday ? HH:MM '))
	sheet['H21'] = tu1
	sheet['J21'] = tu2
	tudelta = datetime.strptime(tu2, FMT) - datetime.strptime(tu1, FMT)
	if tudelta.days < 0:	
		tudelta = timedelta(days=0,seconds=tudelta.seconds, microseconds=tudelta.microseconds)
	sheet['L21'] = tudelta
	
	w1 = (input('What was your start time on Wednesday ? HH:MM '))
	w2 = (input('What was your end time on Wednesday ? HH:MM '))
	sheet['H22'] = w1
	sheet['J22'] = w2
	wdelta = datetime.strptime(w2, FMT) - datetime.strptime(w1, FMT)
	if wdelta.days < 0:	
		wdelta = timedelta(days=0,seconds=wdelta.seconds, microseconds=wdelta.microseconds)
	sheet['L22'] = wdelta
	
	th1 = (input('What was your start time on Thursday ? HH:MM '))
	th2 = (input('What was your end time on Thursday ? HH:MM '))
	sheet['H23'] = th1
	sheet['J23'] = th2
	thdelta = datetime.strptime(th2, FMT) - datetime.strptime(th1, FMT)
	if thdelta.days < 0:	
		thdelta = timedelta(days=0,seconds=thdelta.seconds, microseconds=thdelta.microseconds)
	sheet['L23'] = thdelta
	
	f1 = (input('What was your start time on Friday ? HH:MM '))
	f2 = (input('What was your end time on Friday ? HH:MM '))
	sheet['H24'] = f1
	sheet['J24'] = f2
	fdelta = datetime.strptime(f2, FMT) - datetime.strptime(f1, FMT)
	if fdelta.days < 0:	
		fdelta = timedelta(days=0,seconds=fdelta.seconds, microseconds=fdelta.microseconds)
	sheet['L24'] = fdelta
	
'''
w1 = (input('What was you start time ? HH:MM '))
sheet['H20'] = w1
w2 = (input('What was your end time ? HH:MM '))
sheet['J20'] = w2
tdelta = datetime.strptime(w2, FMT) - datetime.strptime(w1, FMT)
if tdelta.days < 0:
    tdelta = timedelta(days=0,
                seconds=tdelta.seconds, microseconds=tdelta.microseconds)
sheet['L20'] = tdelta
'''

# Work out the date from Mon to Sun
print("")
date_entry = input('Enter the weekending (sunday) in DD-MM-YY format ')

# convert date into datetime object
date1 = datetime.strptime(date_entry, "%d-%m-%y")  

new_date = date1 -timedelta(days=1)  # subtract 1 day from date
new_dateM = date1 -timedelta(days=6)  # subtract 6 day from date
new_dateTu = date1 -timedelta(days=5)  # subtract 5 day from date
new_dateWe = date1 -timedelta(days=4)  # subtract 4 day from date
new_dateTh = date1 -timedelta(days=3)  # subtract 3 day from date
new_dateF = date1 -timedelta(days=2)  # subtract 2 day from date

# convert date into original string like format 
new_date = datetime.strftime(new_date, "%d-%m-%y")
new_dateM = datetime.strftime(new_dateM, "%d-%m-%y")
new_dateTu = datetime.strftime(new_dateTu, "%d-%m-%y")
new_dateWe = datetime.strftime(new_dateWe, "%d-%m-%y")
new_dateTh = datetime.strftime(new_dateTh, "%d-%m-%y")
new_dateF = datetime.strftime(new_dateF, "%d-%m-%y") 

sheet['T3'] = date_entry
sheet['E25'] = date_entry
sheet['E26'] = new_date

sheet['E24'] = new_dateM
sheet['E23'] = new_dateTu
sheet['E22'] = new_dateWe
sheet['E21'] = new_dateTh
sheet['E20'] = new_dateF

#End of program and save
spinner = Spinner()
spinner.start()
print('')
print('')
print('')
print('')
print ('\033[91m' + '\\\Please manually input any overtime breakdown into saved document///' + '\033[91m')
time.sleep(3)
spinner.stop()
wb.save(name + '_timesheet.xlsx')