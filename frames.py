from Tkinter import *
import tkinter.ttk
import tkMessageBox
import itertools
import re
import shutil
import os
import sys
import fileinput
import time
import pandas as pd
from pandas import ExcelFile

# Function to find and replace one specific term in 
# a directory of one or more files
def singleReplace(old_term, new_term, file_type, path_to_directory):
	os.chdir(path_to_directory)
	for root, dir, files in os.walk(path_to_directory):
		for file in files:
			if file.endswith(file_type):
				try:
					f = open(file, 'r')
				except IOError:
					tkMessageBox.showinfo("Error", "File " + str(file) + " failed to open.")
					app.mainloop()
				filedata = f.read()
				f.close()
				if(re.search(str(old_term).encode('string-escape'), filedata, flags=re.IGNORECASE) is not None):
					newdata = re.sub((str(old_term).encode('string-escape')), (str(new_term).encode('string-escape')), filedata, flags=re.IGNORECASE)
					try:
						f = open(file, 'w')
					except IOError:
						tkMessageBox.showinfo("Error", "File " + str(file) + " failed to open.")
						app.mainloop()
					f.write(newdata)
					f.close()
				else:
					pass
			else:
				pass


# Function to find and replace a set of
# terms specified from an Excel workbook and
# sheetname	in a directory of one or more files			
def excelReplace(path_to_excel, column_name_old, column_name_new, shName, file_type, path_to_directory):

	df = pd.read_excel(path_to_excel, sheet_name=shName)
	List_Old_Tags = df[column_name_old]
	List_New_Tags = df[column_name_new].fillna("{[Empty]}")
	sortedZip = sorted(itertools.izip_longest(List_Old_Tags, List_New_Tags, fillvalue="{[Empty]}"), key=lambda oldTag: len(str(oldTag[0])))[::-1]
	os.chdir(path_to_directory)
	# Used to update user on files completed
	x = 0
	for root, dir, files in os.walk(path_to_directory):	
		for file in files:
			if file.endswith(file_type):
					for i, j in sortedZip:
						if str(j) == "{[Empty]}":
							pass
						else:
							try:
								f = open(file, 'r')
							except IOError:
								tkMessageBox.showinfo("Error", "File " + str(file) + " failed to open.")
								app.mainloop()
							filedata = f.read()
							f.close()
							if (re.search('(?i)' + str(i).encode('string-escape'), filedata, flags=re.IGNORECASE) is not None):
								newdata = re.sub((str(i).encode('string-escape')), (str(j).encode('string-escape')), filedata, flags=re.IGNORECASE)
								try:
									f = open(file, 'w')
								except IOError:
									tkMessageBox.showinfo("Error", "File " + str(file) + " failed to open.")
									app.mainloop()
								f.write(newdata)
								f.close()
							else:
								pass
					x = x + 1
			else:
				pass
		else:
			pass

			
# Function to find one specific term in a directory of
# one or more files
def find(searchTerm, file_type, path_to_directory):

	foundList=[]
	os.chdir(path_to_directory)
	for root, dir, files in os.walk(path_to_directory):
		for file in files:
			if file.endswith(file_type):
				try:
					f = open(file, 'r')
				except IOError:
					tkMessageBox.showinfo("Error", "File " + str(file) + " failed to open.")
					app.mainloop()
				filedata = f.read()
				f.close()
				if (re.search(str(searchTerm).encode('string-escape'), filedata, flags=re.IGNORECASE) is not None):
					foundList.append(str(file))
				else:
					pass
			else:
				pass
	#TODO: Give user a prompt with results of search
	if not foundList:
		tkMessageBox.showinfo("Alert", str(searchTerm) + " not found in files")
	else:
		tkMessageBox.showinfo("Results", str(searchTerm) + " Found in files: \n".join(map(" : ", foundList)))
			
	del foundList[:]
	
	
# Function to determine which Find or Find & Replace function to call
# No parameters, variables are initialized inside based on the option selected
# in the app.mainloop()
def determineFunction():
	# Test Function Calls Here
	# Initialize variables based on the option selected
	# Options are: Single Find, Single Replace, or Excel Replace
	if(replaceStatus.get() == "Find"):
		searchTerm = stringFind.get()
		file_type = fileStatus.get()
		path_to_directory = drivePath1.get()
		variableList = [searchTerm, file_type, path_to_directory]
		emptyVariableList = []
		for x in variableList:
			if x == "" or x == " ":
				emptyVariableList.append(x)
		if not emptyVariableList:
			# List is empty
			pass
		else:
			tkMessageBox.showinfo("Error", "One or more required fields is empty.")
			# Delete to avoid memory leaks
			del variableList[:]
			del emptyVariableList[:]
			app.mainloop()
		if(os.path.exists(path_to_directory) is True):
			pass
		else:
			tkMessageBox.showinfo("Error", "Directory path not found. Try again.")
			drivePath1.set("")
			app.mainloop()
		
		# Delete to avoid memory leaks
		del variableList[:]
		del emptyVariableList[:]
		# Call find() if error checking passes
		find(searchTerm, file_type, path_to_directory)
				
	elif(replaceStatus.get() == "SingleReplace"):
		old_term = stringVal.get()
		new_term = stringVal2.get()
		file_type = fileStatus2.get()
		path_to_directory = drivePath2.get()
		
		variableList = [old_term, new_term, file_type, path_to_directory]
		emptyVariableList = []
		for x in variableList:
			if x == "" or x == " ":
				emptyVariableList.append(x)
		if not emptyVariableList:
			# List is empty
			pass
		else:
			tkMessageBox.showinfo("Error", "One or more required fields is empty.")
			# Delete to avoid memory leaks
			del variableList[:]
			del emptyVariableList[:]
			app.mainloop()
		if(os.path.exists(path_to_directory) is True):
			pass
		else:
			tkMessageBox.showinfo("Error", "Directory path not found. Try again.")
			drivePath2.set("")
			app.mainloop()
		
		# Delete to avoid memory leaks
		del variableList[:]
		del emptyVariableList[:]
		# Call singReplace() if error checking passes
		singleReplace(old_term, new_term, file_type, path_to_directory)
		
	elif(replaceStatus.get() == "ExcelReplace"):
		path_to_excel = workbookEntryText.get()
		column_old_name = oldColumnName.get()
		column_new_name = newColumnName.get()
		shName = sheetName.get()
		file_type = fileStatus3.get()
		path_to_directory = drivePath3.get()
		
		variableList = [path_to_excel, column_old_name, column_new_name, shName, file_type, path_to_directory]
		emptyVariableList = []
		for x in variableList:
			if x == "" or x == " ":
				emptyVariableList.append(x)
		if not emptyVariableList:
			# List is empty
			pass
		else:
			tkMessageBox.showinfo("Error", "One or more required fields is empty.")
			# Delete to avoid memory leaks
			del variableList[:]
			del emptyVariableList[:]
			app.mainloop()
		if(os.path.exists(path_to_directory) is True and os.path.exists(path_to_excel) is True):
			pass
		elif(os.path.exists(path_to_directory) is True and os.path.exists(path_to_excel) is not True):
			tkMessageBox.showinfo("Error", "Excel path does not exist. Try again.")
			workbookEntryText.set("")
			app.mainloop()
		elif(os.path.exists(path_to_directory) is not True and os.path.exists(path_to_excel) is True):
			tkMessageBox.showinfo("Error", "Directory path does not exist. Try again.")
			drivePath3.set("")
			app.mainloop()
		else:
			tkMessageBox.showinfo("Error", "Neither directory nor Excel path found. Try again.")
			drivePath3.set("")
			workbookEntryText.set("")
			app.mainloop()
		
		# Delete to avoid memory leaks
		del variableList[:]
		del emptyVariableList[:]
		# Call excelReplace() if error checking passes
		excelReplace(path_to_excel, column_old_name, column_new_name, shName, file_type, path_to_directory)
	else:
		print "No function specified, please select from the top 3 options."
		
		
def reset():
	replaceStatus.set(None)
	drivePath1.set("")
	stringFind.set("")
	fileStatus1.set("")
	drivePath2.set("")
	stringVal.set("")
	stringVal2.set("")
	fileStatus2.set("")
	drivePath3.set("")
	workbookEntryText.set("")
	sheetName.set("")
	oldColumnName.set("")
	newColumnName.set("")
	fileStatus3.set("")
	app.mainloop()
		


app = Tk()
app.title("Find & Replace: Saginaw Power & Automation")
app.geometry("925x400")

topLeftFrame = Frame(app, width=300, height=250, highlightcolor="black")
topLeftFrame.grid(row=0, column=0)
topMiddleFrame = Frame(app, width=300, height=250)
topMiddleFrame.grid(row=0, column=1)

topRightFrame = Frame(app, width=300, height=250)
topRightFrame.grid(row=0, column=2)

leftFrame = Frame(app, width=300, height=250)
leftFrame.grid(row=1, column=0)

middleFrame = Frame(app, width=300, height = 250)
middleFrame.grid(row=1, column=1, padx=50, pady=0)

rightFrame = Frame(app, width=300, height=250)
rightFrame.grid(row=1, column=2, padx=0, pady=25)

buttonFrame = Frame(app, width=300, height=250)
buttonFrame.grid(row=2, column=1)

#######Left Middle Frame	
replaceStatus = StringVar()
replaceStatus.set(None)
radio10 = Radiobutton(topLeftFrame, text="Single Find", value="Find", variable=replaceStatus).pack()

driveText1 = StringVar()
driveText1.set("Enter Path to Directory:")
driveLabel1 = Label(leftFrame, textvariable=driveText1).pack()
drivePath1 = StringVar()
driveEntry1 = Entry(leftFrame, textvariable=drivePath1).pack()

stringText1 = StringVar()
stringText1.set("Enter string to find: ")
stringLabel1 = Label(leftFrame, textvariable=stringText1).pack()
stringFind = StringVar()
stringFindEntry = Entry(leftFrame, textvariable=stringFind).pack()
	
fileText1 = StringVar()
fileText1.set("Enter file ext of files to be searched (.xml, .txt, etc.): ")
fileLabel1 = Label(leftFrame, textvariable=fileText1).pack()
fileStatus1 = StringVar()
fileStatusEntry1 = Entry(leftFrame, textvariable=fileStatus1).pack()
#######Left Middle Frame
#######Middle Frame
radio10 = Radiobutton(topMiddleFrame, text="Single Replace", value="SingleReplace", variable=replaceStatus).pack()

driveText2 = StringVar()
driveText2.set("Enter Path to Directory:")
driveLabel2 = Label(middleFrame, textvariable=driveText2).pack()
drivePath2 = StringVar()
driveEntry2 = Entry(middleFrame, textvariable=drivePath2).pack()

stringText2 = StringVar()
stringText2.set("Enter string to find: ")
stringLabel2 = Label(middleFrame, textvariable=stringText2).pack()
stringVal = StringVar()
stringEntry1 = Entry(middleFrame, textvariable=stringVal).pack()

stringText3 = StringVar()
stringText3.set("Enter replacement string: ")
stringLabel3 = Label(middleFrame, textvariable=stringText3).pack()
stringVal2 = StringVar()
stringEntry2 = Entry(middleFrame, textvariable=stringVal2).pack()

fileText2 = StringVar()
fileText2.set("Enter file ext of files to be searched (.xml, .txt, etc.): ")
fileLabel2 = Label(middleFrame, textvariable=fileText2).pack()
fileStatus2 = StringVar()
fileStatusEntry2 = Entry(middleFrame, textvariable=fileStatus2).pack()
#######Middle Frame
#######Right Middle Frame
radio10 = Radiobutton(topRightFrame, text="Excel Replace", value="ExcelReplace", variable=replaceStatus).pack()

driveText3 = StringVar()
driveText3.set("Enter Path to Directory:")
driveLabel3 = Label(rightFrame, textvariable=driveText3).pack()
drivePath3 = StringVar()
driveEntry3 = Entry(rightFrame, textvariable=drivePath3).pack()

workbookText = StringVar()
workbookText.set("Enter 'Path\to\excelworkbook.xlsx':")
workbookLabel = Label(rightFrame, textvariable=workbookText).pack()
workbookEntryText = StringVar()
workbookEntry = Entry(rightFrame, textvariable=workbookEntryText).pack()

sheetText = StringVar()
sheetText.set("Enter sheet name inside workbook: ")
sheetLabel = Label(rightFrame, textvariable=sheetText).pack()
sheetName = StringVar()
sheetEntry = Entry(rightFrame, textvariable=sheetName).pack()

oldColumnText = StringVar()
oldColumnText.set("Enter the name of the column of old terms:")
oldColumnLabel = Label(rightFrame, textvariable=oldColumnText).pack()

oldColumnName = StringVar()
oldColumnEntry = Entry(rightFrame, textvariable=oldColumnName).pack()

newColumnText = StringVar()
newColumnText.set("Enter the name of the column of new terms:")
newColumnLabel = Label(rightFrame, textvariable=newColumnText).pack()

newColumnName = StringVar()
newColumnEntry = Entry(rightFrame, textvariable=newColumnName).pack()

fileText3 = StringVar()
fileText3.set("Enter file ext of files to be searched (.xml, .txt, etc.): ")
fileLabel3 = Label(rightFrame, textvariable=fileText3).pack()
fileStatus3 = StringVar()
fileStatusEntry3 = Entry(rightFrame, textvariable=fileStatus3).pack()
#######Right Middle Frame
#######Button Frame
button1 = Button(buttonFrame, text="Submit", width=20, command=determineFunction).grid(row=0, column=0, padx=0, pady=25)
button2 = Button(buttonFrame, text="Reset", width=20, command=reset).grid(row=0, column=1, padx=0, pady=25)
#######Button Frame
# Create lines to separate different functions of app
tkinter.ttk.Separator(app, orient=VERTICAL).grid(column=0, row=1, rowspan=1, sticky='nse')
tkinter.ttk.Separator(app, orient=VERTICAL).grid(column=1, row=1, rowspan=1, sticky='nse')
tkinter.ttk.Separator(app, orient=HORIZONTAL).grid(column=0, row=1, columnspan=3, sticky='wen')
tkinter.ttk.Separator(app, orient=HORIZONTAL).grid(column=0, row=2, columnspan=3, sticky='wen')
	
app.mainloop()
