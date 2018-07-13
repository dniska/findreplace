from Tkinter import *
import tkMessageBox
import itertools
import re
import shutil
import os
import sys
import fileinput
import time
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

app = Tk()
app.title("Find & Replace: Saginaw Power & Automation")
app.geometry("500x400")


leftFrame = Frame(app, width=300, height=250)
leftFrame.grid(row=0, column=0)
#leftFrame.pack()

leftBottomFrame = Frame(app, width=300, height=250)
leftBottomFrame.grid(row=1, column=0)
#leftBottomFrame.pack()

rightFrame = Frame(app, width=300, height=250)
rightFrame.grid(row=0, column=1)
#rightFrame.pack()

rightBottomFrame = Frame(app, width=300, height=250)
rightBottomFrame.grid(row=1, column=1)
#rightBottomFrame.pack()

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
	
	
def excelReplace(path_to_excel, column_name_old, column_name_new, shName, file_type, path_to_directory):

	df = pd.read_excel(path_to_excel, sheet_name=sheetName)
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
def determineFunction():
	# Test Function Calls Here
	if(findStatus.get() == "Replace" and repStatus.get() == "Excel"):
		print "excelReplace() called"
		
		excel_path = workbookName.get()
		old_column = column1Name.get()
		new_column = column2Name.get()
		shName = sheetName.get()
		file_type = fileStatus.get()
		directory_path = drivePath.get()
		variableList = [excel_path, old_column, new_column, shName, file_type, directory_path]
		emptyList = []
		for x in variableList:
			if x == "":
				emptyList.append(x)
		tkMessageBox.showinfo("Error", "The following fields are still empty: \n".join(map(" : ", emptyList)))
		if not emptyList:
			del emptyList[:]
			app.mainloop()
		if(os.path.exists(directory_path) is True and os.path.exists(excel_path) is True):
			pass
		elif(os.path.exists(directory_path) is not True and os.path.exists(excel_path) is True):
			tkMessageBox.showinfo("Error", "Directory path does not exist. Try again.")
			drivePath.set("")
			app.mainloop()
		elif(os.path.exists(directory_path) is True and os.path.exists(excel_path) is not True):
			tkMessageBox.showinfo("Error", "Excel workbook does not exist or file path is invalid. Try again.")
			workbookName.set("")
			app.mainloop()
		else:
			tkMessageBox.showinfo("Error", "Both directory path and excel path are invalid. Try again.")
			drivePath.set("")
			workbookName.set("")
			app.mainloop()
			#excelReplace(excel_path, old_column, new_column, shName, file_type, directory_path)
		
		
		# excelReplace Currently works
		# excelReplace(r'C:\\Users\\Daniel.Niska\\PythonScripts\\PyProject\\ExcelTestDir\\TestBook.xlsx', 'Column 1', 'Column 2', 'test_sheet', '.xml', r'C:\\Users\\Daniel.Niska\\PythonScripts\\PyProject\\ExcelTestDir')
		# Raise Exception for incorrect input
		
	elif(findStatus.get() == "Find" and (repStatus.get() == "None")):
		print "find() called"
		searchTerm = findString.get()
		file_type = fileStatus.get()
		directory_path = drivePath.get()
		variableList = [searchTerm, file_type, directory_path]
		emptyList = []
		for x in variableList:
			if x == "":
				emptyList.append(x)
		tkMessageBox.showinfo("Error", "The following fields are still empty: \n".join(map(" : ", emptyList)))
		if not emptyList:
			del emptyList[:]
			app.mainloop()
			#find(searchTerm, file_type, directory_path)
		while(os.path.exists(directory_path) is not True):
			tkMessageBox.showinfo("Error", "Directory does not exist")
			drivePath.set("")
			app.mainloop()
		
		
		# find Currently Works
		# find('NewTag', '.xml', r'C:\\Users\\Daniel.Niska\\PythonScripts\\PyProject\\ExcelTestDir')
		# Raise Exception for incorrect input
	
	elif(repStatus.get() == "Single" and findStatus.get() == "Replace"):
		print "singleReplace called"
		old_term = findString.get()
		new_term = replaceString.get()
		file_type = fileStatus.get()
		directory_path = drivePath.get()
		variableList = [old_term, new_term, file_type, directory_path]
		emptyList = []
		for x in variableList:
			if x == "":
				emptyList.append(x)
		tkMessageBox.showinfo("Error", "The following fields are still empty: \n".join(map(" : ", emptyList)))
		if not emptyList:
			del emptyList[:]
			app.mainloop()
			
		#singleReplace(old_term, new_term, file_type, directory_path)
		
		
		# singleReplace Currently Works
		# singleReplace('NewTag', 'replaced', '.xml', r'C:\\Users\\Daniel.Niska\\PythonScripts\\PyProject\\ExcelTestDir')
		# Raise Exception for incorrect input
	else:
		print "Exception Raised"
		# Raise Exception: Incorrect sequence of selections
		
def reset():
	driveStatus.set(None)
	repStatus.set(None)
	findStatus.set(None)
	findString.set("")
	replaceString.set("")
	fileStatus.set("")
	drivePath.set("")
	column1Name.set("")
	column2Name.set("")
	workbookName.set("")
	sheetName.set("")
	app.mainloop()
		
############################## Left Frame #####################################

driveStatus = StringVar()
driveStatus.set(None)
radio1 = Radiobutton(leftFrame, text="Network", value="Network", variable=driveStatus).pack()
radio1 = Radiobutton(leftFrame, text="Local", value="Local", variable=driveStatus).pack()

driveText = StringVar()
driveText.set("Enter path to directory: ")
driveLabel = Label(leftFrame, textvariable=driveText).pack()

# Text Field for Drive Name
drivePath = StringVar()
driveEntry = Entry(leftFrame, textvariable=drivePath).pack()

############################## Left Frame #####################################

############################## Left Bottom Frame #####################################

# Radio Buttons: Find or Find & Replace
	# Text field for old string to find
	# Text field for new string to replace
findStatus = StringVar()
findStatus.set(None)
radio3 = Radiobutton(leftBottomFrame, text="Find", value="Find", variable=findStatus).pack()
radio3 = Radiobutton(leftBottomFrame, text="Find & Replace", value="Replace", variable=findStatus).pack()

repStatus = StringVar()
repStatus.set(None)
radio2 = Radiobutton(leftBottomFrame, text="Single Replacement", value="Single", variable=repStatus).pack()
radio2 = Radiobutton(leftBottomFrame, text="Replacement From Excel", value="Excel", variable=repStatus).pack()

stringText = StringVar()
stringText.set("Enter string to replace or find: ")
stringLabel = Label(leftBottomFrame, textvariable=stringText).pack()
#stringLabel.grid(row=0, column=0)
findString = StringVar()
findEntry = Entry(leftBottomFrame, textvariable=findString).pack()
#findEntry.grid(row=0, column=1)

replaceText = StringVar()
replaceText.set("Enter new String (if applicable): ")
replaceLabel = Label(leftBottomFrame, textvariable=replaceText).pack()
#replaceLabel.grid(row=1, column=0)
replaceString = StringVar()
replaceEntry = Entry(leftBottomFrame, textvariable=replaceString).pack()
#replaceEntry.grid(row=1, column=1)

############################## Left Bottom Frame #####################################

############################## Right Frame #####################################

# Text for Excel workbook name, sheet name, and column names
wbText = StringVar()
wbText.set("Enter path to excel workbook (if applicable): ")
wbLabel = Label(rightFrame, textvariable=wbText).pack()
workbookName = StringVar()
workbookEntry = Entry(rightFrame, textvariable=workbookName).pack()

sheetText = StringVar()
sheetText.set("Enter sheet name inside workbook (if applicable): ")
sheetLabel = Label(rightFrame, textvariable=sheetText).pack()
sheetName = StringVar()
sheetEntry = Entry(rightFrame, textvariable=sheetName).pack()

column1Text = StringVar()
column1Text.set("Enter column name of old strings (if applicable): ")
column1Label = Label(rightFrame, textvariable=column1Text).pack()
column1Name = StringVar()
column1Entry = Entry(rightFrame, textvariable=column1Name).pack()

column2Text = StringVar()
column2Text.set("Enter column name of new strings (if applicable): ")
column2Label = Label(rightFrame, textvariable=column2Text).pack()
column2Name = StringVar()
column2Entry = Entry(rightFrame, textvariable=column2Name).pack()

fileText = StringVar()
fileText.set("Enter file ext of files to be searched (.xml, .txt, etc.): ")
fileLabel = Label(rightFrame, textvariable=fileText).pack()
fileStatus = StringVar()
fileStatusEntry = Entry(rightFrame, textvariable=fileStatus).pack()
############################## Right Frame #####################################

############################## Right Bottom Frame #####################################

# Submit Button
	# Submit button calls script with find & replace functionality
	# Create a class, passing in user input
button1 = Button(rightBottomFrame, text="Submit", width=20, command=determineFunction).pack()
button2 = Button(rightBottomFrame, text="Reset", width=20, command=reset).pack()

############################## Right Bottom Frame #####################################


app.mainloop()