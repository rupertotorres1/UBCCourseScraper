import xlrd

# Use the excel workbook with the courses
book = xlrd.open_workbook("UBCCourseSubjects.xlsx")

courseList = []

# Read each sheet (each sheet is for a subject)
for sheet in book.sheets():
	startRow = 0
	endRow   = 0
	noCourses = False

	# Mark the row where the offered courses would start to be listed in the sheet
	for row in range(sheet.nrows):
		if (sheet.cell(row,0).value == "Course"):
			startRow = row + 1
			# Or whether no courses are offered
		if (sheet.cell(row,0).value == "No courses offered for 2017 Winter."):
			noCourses = True

	# Only look for courses if the subject actually offers them
	if (not(noCourses)):
		for row in range(startRow, sheet.nrows):
			if (sheet.cell(row,0).value == "Browse"):
				endRow = row - 1

		# put in a list all of the courses offered at ubc (using the excel document)
		for row in range(startRow, endRow):
			course = sheet.cell(row, 0).value
			courseList.append(course)
			


# Now we are going to the courseList to look up pre-requisites and co-requisites in each course's website 
from bs4 import BeautifulSoup as soup
from urllib.request import urlopen as uReq
import time

dict = {}


for c in courseList:
	
	dept      = c.split()[0]
	courseNum = c.split()[1]

	myUrl = "https://courses.students.ubc.ca/cs/main?pname=subjarea&tname=subjareas&req=3&dept=" + dept + "&course=" + courseNum

	# opening up connection
	uClient = uReq(myUrl)
	pageHtml = uClient.read()
	uClient.close()

	# html parsing
	pageSoup = soup(pageHtml, "html.parser")

	# Find all the p tags in the html (that's we find the pre-reqs and co-reqs)
	pTags = pageSoup.findAll("p")

	preReqText = ""
	preReqs    = []
	coReqText  = ""
	coReqs     = []
	preCoReqs  = []

	for pTag in pTags:
		if ("Pre-reqs:" in str(pTag)):
			preReqText = str(pTag)
		if ("Co-reqs:" in str(pTag)):
			coReqText = str(pTag)

	# Pre-reqs and co-reqs separated for maybe printing them separately in a future updated
	for x in courseList:
		if (x in preReqText):
				preReqs.append(x)
		if (x in coReqText):
				coReqs.append(x)

	preCoReqs.extend(preReqs)
	preCoReqs.extend(coReqs)

	# Key-value pairs of courses and their list of pre-requisites and co-requisites
	dict[c] = preCoReqs

# pickle the dict to use it in another python file (so we don't have to use the excel workbook and then look up each website every time)
import pickle
pickle.dump(dict, open("dictCoursesPreCoReqs.p", "wb"), protocol=2)
