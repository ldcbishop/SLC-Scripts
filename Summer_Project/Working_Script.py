"""
This is a small script to migrate Excel files with on format to a format acceptable to the Filemaker system.

This file is being made during May of 2015 at the Behest of the Sanger Learning Center.

Primary Coder: Logan D.C. Bishop

Any questions/concerns about this code can be directed to ldcbishop@gmail.com

Working version of openpyxl requires python 2.2 usage.  Can be run with "python" command.

Not robust for widely varied excel files 
"""


import sys #here to allow argument passing
import os #included to splite filename from extension
import re #regex package
from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl import load_workbook


"""
A student class which stores all variables in reference to a single student
"""
class Student:
	name = ""
	unique_num = 0
	uteid = ""

	def __init__(StuName, StuUN, Stu_uteid):
		self.name = StuName
		self.unique_num = StuUN
		self.uteid = Stu_uteid

	def returnSelf():
		return (self.name, self.unique_num, self.uteid)

"""
A function to find a key word in an excel spreadsheet and return its coordinate
@param: worksheet, and a valid string to look for.
@return: a coordinate string
"""
def searchWorkBook(work_sheet, search_string):
	#Assume Unique values are in the first 30 columns and first 50 rows
	search_space = list(work_sheet.iter_rows('A1:M7'))
	cell_re = re.compile('[A-Z]+[0-9]+')
	for x in search_space:
		for y in x:
			if str(y.value).lower() == search_string.lower():
				return cell_re.findall(str(y))[0]
	return '00'

"""
A function to take the string excel coordinate and translate it into numerical coordinates
@param: cell, a coordinate string
@return: tuple with numerical coordinates (x, y)
"""
def translateCoord(cell):
	#Take a string, return a numerical tuple
	alphabetConvString = 64
	row = cell[1:]
	column = ord(cell[0]) - alphabetConvString 
	return (int(row), int(column))

"""
Take a date in the origianl format e.g. 9_16 and return a properly formatted date_origin
@param: bad_date, a badly formatted date string
@return: good_date, a properly formatted date string 
"""
def parseAndBuildDate(bad_date):
	date_re = re.compile('[1-9]+\/[0-9]+')
	month_day = re.search(date_re, bad_date).group(0).split('/')
	return month_day[0]+'/'+month_day[1]+'/'

"""
System Main Method
"""

"""
Command format: python Working_Script.py myWork_book.xlsx
"""

src_filename = sys.argv[1]
wb_origin = load_workbook(src_filename, data_only=True) #First argument should be script names
wb_output = Workbook() #Create destination workbook
	#sheet_output = wb_output.active() #pulls active worksheet to pool data
	#dest_filename = os.path.splitext(src_filename)[0]+"_parsed.xlsx" #Name that the new workbook will be saved under
#Work book relations are now 1:1 with each sheet being 1:1

for sheet in wb_origin :

	active_sheet = wb_output.create_sheet()

	
	semester_year = sheet['A1'].value
	year_re = re.compile('20[0-9][0-9]')
	year = year_re.findall(semester_year)[0]

	#Two cases, either fall or Spring
	fall_re = re.compile('[Ff]+[Aa][Ll][Ll]')
	semester = fall_re.search(semester_year)
	if str(type(semester)) == "<type '_sre.SRE_Match'>" :
		CCYYS = year + '8'
	else :
		CCYYS = year + '2'


	#Grabs list of all unique numbers in the spreadsheet
	un_string = sheet['A2'].value
	un_re = re.compile('[0-9][0-9][0-9][0-9][0-9]')
	unique_Num_list = un_re.findall(str(un_string))

	#Where the word date is located
	date_origin = translateCoord(searchWorkBook(sheet, 'date:'))
	#Name is leftmost cell, followed by EID and Unique #
	name_origin = translateCoord(searchWorkBook(sheet, 'student'))

	ex = parseAndBuildDate("T 9/16") 

	print ex + year
	"""
	print "My searchable index is " + searchWorkBook(sheet, 'date:')
	print "The value I found at this index " + str(sheet.cell(row = date_origin[0], column = date_origin[1]).value)
	print "The index I am reporting " + str(date_origin)
	"""

"""
Current print format
ccyys, eid, name, unique, date

Getting a cell:
cell = sheet.cell(row = x, column = y)

Setting a value:
cell.value = "hello, world"
"""



