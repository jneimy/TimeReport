from __future__ import division
from collections import OrderedDict
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from datetime import date, timedelta, datetime
import os
import sys
import requests
import json
import re

excel_filename = "Timesheet_Report.xlsx"

harvest_headers = {
	'Content-type': 'application/json',
	'Accept': 'application/json',
	'Authorization': 'Basic Y21vcmlrdW5pQHJldmFjb21tLmNvbToqNjAlaEZ4ViVSWHU='
}

hoursForAcceptance = 8

# Excel Style
namesStart = (4, 1)
totCell = (1, 7)
consultantsFont = Font(name='Calibri',
							size=12,
							color='00000000')
text = Font(name='Calibri',
					size=12)

def init():
	# Harvest - Build request & load projects
	projects = requests.get('https://revacomm.harvestapp.com/projects', headers=harvest_headers)
	projects_json = projects.json()

	return projects_json


def peopleTime(date):
	fteTime = {}
	contTime = {}
	people = requests.get('https://revacomm.harvestapp.com/people', headers=harvest_headers)
	people_json = people.json()
	for person in people_json:
		pUser = person['user']
		uid = str(pUser['id'])
		first = pUser['first_name']
		last = pUser['last_name']

		# Skip old employees
		if not pUser['is_active']:
			continue

		# Identify contractors
		isContract = pUser['is_contractor']

		userTime = requests.get('https://revacomm.harvestapp.com/people/' + uid + '/entries?from=' + yesterday + '&to=' + yesterday, headers=harvest_headers)
		userTime_json = userTime.json()

		marked = False
		enteredTime = 0
		if userTime_json is not None:
			for ut in userTime_json:
				entry = ut.get("day_entry", 0)
				if entry == 0:
					continue
				hours = entry.get("hours", 0)
				enteredTime += hours
			if enteredTime >= hoursForAcceptance:
				marked = True

		if isContract:
			contTime[first + " " + last] = marked
		else:
			fteTime[first + " " + last] = marked
	return (fteTime, contTime)


def openExcel(filename, weekSheetName, fteTime, contTime):
	wb = None
	ws = None

	# Open or create a new worksheet
	isCreateWs = False
	if os.path.exists(filename):
		wb = load_workbook(filename)
		if weekSheetName in wb.get_sheet_names():
			ws = wb[weekSheetName]
		else:
			isCreateWs = True
	else:
		wb = Workbook()
		ws = wb.active

		# Delete current sheet & create new
		if ws is not None:
			wb.remove_sheet(ws)
		isCreateWs = True

	if isCreateWs:
		wsSrc = wb["Template"]
		ws = wb.copy_worksheet(wsSrc)
		ws.title = weekSheetName
		ws.cell(row=1, column=1, value=weekSheetName)

		# Reorder
		sheetInd = len(wb.get_sheet_names()) - 1
		wb._sheets = [wb._sheets[sheetInd]] + wb._sheets[0:sheetInd]
		ws.active = 0

		# Setup Template
		row = namesStart[0]
		for key in sorted(fteTime.iterkeys()):
			ws.cell(row=row, column=1, value=key)
			row += 1

		# Consultant Header
		row += 1
		ws.cell(row=row, column=1, value="Consultants")
		row += 1

		for key in sorted(contTime.iterkeys()):
			ws.cell(row=row, column=1, value=key)

		# Setup Formulas
	return (wb, ws)


def closeExcel(wb, filename):
	wb.save(filename)


# TODO: make it dynamic by adding in formulas
def dynamicOutputToExcel(ws, dayOfWeek, fteTime, contTime):
	# 2 for one space and base 1
	dayToCol = dayOfWeek + 2

	row = namesStart[0]
	for key in sorted(fteTime.iterkeys()):
		cell = ws.cell(row=row, column=1)
		if key == cell.value:
			ws.cell(row=row, column=dayToCol, value=int(fteTime[key]))
		else:
			print "ERROR: invalid person key: " + key + " cell: " + cell.value
		row += 1

	# increase row to start of contractors
	row += 2
	for key in sorted(contTime.iterkeys()):
		cell = ws.cell(row=row, column=1)
		if key == cell.value:
			ws.cell(row=row, column=dayToCol, value=int(contTime[key]))
		else:
			print "ERROR: invalid person key: " + key + " cell: " + cell.value
		row += 1


def outputToExcel(ws, dayOfWeek, fteTime, contTime):
	# 2 for one space and base 1
	dayToCol = dayOfWeek + 2

	row = namesStart[0]
	for key in sorted(fteTime.iterkeys()):
		cell = ws.cell(row=row, column=1)
		if key == cell.value:
			ws.cell(row=row, column=dayToCol, value=int(fteTime[key]))
		else:
			print "ERROR: invalid person key: " + key + " cell: " + cell.value
		row += 1

	# increase row to start of contractors
	row += 2
	for key in sorted(contTime.iterkeys()):
		cell = ws.cell(row=row, column=1)
		if key == cell.value:
			ws.cell(row=row, column=dayToCol, value=int(contTime[key]))
		else:
			print "ERROR: invalid person key: " + key + " cell: " + cell.value
		row += 1


if __name__ == '__main__':
	## CI runs this at 8am HNL ##
	setDt = None
	if len(sys.argv) >= 2:
		try:
			setDt = datetime.strptime(sys.argv[1], "%m.%d.%Y")
		except:
			setDt = None

	if setDt:
		yesterdayDt = setDt
	else:
		yesterdayDt = date.today() - timedelta(1)
	yesterday = str(yesterdayDt.strftime('%Y%m%d'))
	yesterdayFmt = str(yesterdayDt.strftime('%m.%d.%Y'))
	print "Time Reporting for: " + yesterdayFmt

	# # M=0, T=1, W=2, T=3, F=4, S=5, S=6
	dayOfWeek = yesterdayDt.weekday()
	startOfWeekDt = yesterdayDt - timedelta(dayOfWeek)
	startOfWeekFmt = str(startOfWeekDt.strftime('%m.%d.%Y'))

	(fteTime, contTime) = peopleTime(yesterday)

	wb, ws = openExcel(excel_filename, startOfWeekFmt, fteTime, contTime)
	dynamicOutputToExcel(ws, dayOfWeek, fteTime, contTime)
	closeExcel(wb, excel_filename)
