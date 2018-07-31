from Tkinter import *
import tkinter.ttk
import tkMessageBox
import fnmatch
import itertools
import re
import os
import sys
import fileinput
import pandas as pd
import Queue
from pandas import ExcelFile
import threading
import time
import random
from threading import Thread

class GuiPart:
    def singleReplace(old_term, new_term, file_type, path_to_directory):
		frame_pb.start()
		os.chdir(path_to_directory.encode('string-escape'))
		foundList=[]
		for root, dir, files in os.walk(path_to_directory.encode('string-escape')):
			for file in files:
				if file.endswith(file_type):
					try:
						f = open(file, 'r')
						filedata = f.read()
						f.close()
					except IOError:
						tkMessageBox.showinfo("Error", "File " + str(file) + " failed to open.")
						frame_pb.stop()
						app.mainloop()
					if(re.search(str(old_term).encode('string-escape'), filedata, flags=re.IGNORECASE) is not None):
						newdata = re.sub((str(old_term).encode('string-escape')), (str(new_term).encode('string-escape')), filedata, flags=re.IGNORECASE)
						foundList.append(str(file))
						try:
							f = open(file, 'w')
							f.write(newdata)
							f.close()
						except IOError:
							tkMessageBox.showinfo("Error", "File " + str(file) + " failed to open.")
							frame_pb.stop()
							app.mainloop()
						try:
							f = open('single_log_file.txt', 'a')
							f.write(str(old_term) + " replaced in file " + str(file) + "\r\n")
							f.close()
						except IOError:
							tkMessageBox.showinfo("Error", "File 'single_log_file.txt' failed to open.")
							pass
					else:
						pass
				else:
					pass
		if not foundList:
			tkMessageBox.showinfo("Alert", "Search term not found in files")
		else:
			tkMessageBox.showinfo("Results", "See file 'single_log_file.txt' for results.")
		frame_pb.stop()
		del foundList[:]

	# Function to find and replace a set of
	# terms specified from an Excel workbook and
	# sheetname	in a directory of one or more files			
    def excelReplace(path_to_excel, column_name_old, column_name_new, shName, file_type, path_to_directory):
		frame_pb.start()
		df = pd.read_excel(path_to_excel.encode('string-escape'), sheet_name=shName)
		List_Old_Tags = df[column_name_old]
		List_New_Tags = df[column_name_new].fillna("{[Empty]}")
		sortedZip = sorted(itertools.izip_longest(List_Old_Tags, List_New_Tags, fillvalue="{[Empty]}"), key=lambda oldTag: len(str(oldTag[0])))[::-1]
		os.chdir(path_to_directory.encode('string-escape'))
		# Used to update user on files completed
		foundList=[]
		for root, dir, files in os.walk(path_to_directory.encode('string-escape')):	
			for file in files:
				if file.endswith(file_type):
						for i, j in sortedZip:
							if str(j) == "{[Empty]}":
								pass
							else:
								try:
									f = open(file, 'r')
									filedata = f.read()
									f.close()
								except IOError:
									tkMessageBox.showinfo("Error", "File " + str(file) + " failed to open.")
									frame_pb.stop()
									app.mainloop()
								if (re.search('(?i)' + str(i).encode('string-escape'), filedata, flags=re.IGNORECASE) is not None):
									foundList.append(str(i))
									newdata = re.sub((str(i).encode('string-escape')), (str(j).encode('string-escape')), filedata, flags=re.IGNORECASE)
									try:
										f = open(file, 'w')
										f.write(newdata)
										f.close()
									except IOError:
										tkMessageBox.showinfo("Error", "File " + str(file) + " failed to open.")
										frame_pb.stop()
										app.mainloop()
									try:
										f = open('excel_log_file.txt', 'a')
										f.write(str(i) + " replaced in file " + str(file) + " with " + str(j) + "\r\n")
										f.close()
									except IOError:
										tkMessageBox.showinfo("Error", "File 'excel_log_file.txt' failed to open.")
										pass
								else:
									pass
				else:
					pass
			else:
				pass
		if not foundList:
			tkMessageBox.showinfo("Alert", "Search terms from Excel data not found in files")
		else:
			tkMessageBox.showinfo("Results", "See file 'excel_log_file.txt' for results.")
		frame_pb.stop()
		del foundList[:]
		
	# Function to find one specific term in a directory of
	# one or more files
    def find(searchTerm, file_type, path_to_directory):
		frame_pb.start()
		foundList=[]
		os.chdir(path_to_directory.encode('string-escape'))
		for root, dir, files in os.walk(path_to_directory.encode('string-escape')):
			for file in files:
				if file.endswith(file_type):
					try:
						f = open(file, 'r')
						filedata = f.read()
						f.close()
					except IOError:
						tkMessageBox.showinfo("Error", "File " + str(file) + " failed to open.")
						frame_pb.stop()
						app.mainloop()
					if (re.search(str(searchTerm).encode('string-escape'), filedata, flags=re.IGNORECASE) is not None):
						foundList.append(str(file))
						try:
							f = open('find_log_file.txt', 'a')
							f.write(str(searchTerm) + " found in file " + str(file) + "\r\n")
							f.close()
						except IOError:
							tkMessageBox.showinfo("Error", "File 'find_log_file.txt' failed to open.")
							pass
					else:
						pass
				else:
					pass
		#TODO: Give user a prompt with results of search
		if not foundList:
			tkMessageBox.showinfo("Alert", str(searchTerm) + " not found in files")
		else:
			tkMessageBox.showinfo("Results", "See file 'find_log_file.txt' for results.")
		frame_pb.stop()		
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
			file_type = fileStatus1.get()
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
			tkMessageBox.showinfo("Error", "No function specified, please select from the top 3 options.")
			app.mainloop()
			
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
		
    def processIncoming(self):
        """Handle all messages currently in the queue, if any."""
        while self.queue.qsize(  ):
            try:
                msg = self.queue.get(0)
                # Check contents of message and do whatever is needed. As a
                # simple test, print it (in real life, you would
                # suitably update the GUI's display in a richer fashion).
                print msg
            except Queue.Empty:
                # just on general principles, although we don't
                # expect this branch to be taken in this case
                pass

    def __init__(self, master, queue, endCommand):
		self.queue = queue
		app.title("Find & Replace: Saginaw Power & Automation")
		app.geometry("925x450")
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
		bottomFrame = Frame(app, width=300, height=250)
		bottomFrame.grid(row=3, column=1)	
		replaceStatus = StringVar()
		replaceStatus.set(None)
		radio10 = Radiobutton(topLeftFrame, text="Single Find", value="Find", variable=replaceStatus).pack()
		driveText1 = StringVar()
		driveText1.set("Enter Path\\to\\Directory:")
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
		radio10 = Radiobutton(topMiddleFrame, text="Single Replace", value="SingleReplace", variable=replaceStatus).pack()
		driveText2 = StringVar()
		driveText2.set("Enter Path\\to\\Directory:")
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
		radio10 = Radiobutton(topRightFrame, text="Excel Replace", value="ExcelReplace", variable=replaceStatus).pack()
		driveText3 = StringVar()
		driveText3.set("Enter Path\\to\\Directory:")
		driveLabel3 = Label(rightFrame, textvariable=driveText3).pack()
		drivePath3 = StringVar()
		driveEntry3 = Entry(rightFrame, textvariable=drivePath3).pack()
		workbookText = StringVar()
		workbookText.set("Enter 'Path\\to\\excelworkbook.xlsx':")
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
		button1 = Button(buttonFrame, text="Submit", width=20, command=GuiPart.determineFunction).grid(row=0, column=0, padx=0, pady=25)
		button2 = Button(buttonFrame, text="Reset", width=20, command=GuiPart.reset).grid(row=0, column=1, padx=0, pady=25)
		pbText = StringVar()
		pbText.set("If this is moving, your task is in-progress")
		pbLabel = Label(bottomFrame, textvariable=pbText).pack()
		frame_pb = tkinter.ttk.Progressbar(bottomFrame, orient='horizontal', mode='indeterminate', maximum=100)
		frame_pb.pack()
		tkinter.ttk.Separator(app, orient=VERTICAL).grid(column=0, row=1, rowspan=1, sticky='nse')
		tkinter.ttk.Separator(app, orient=VERTICAL).grid(column=1, row=1, rowspan=1, sticky='nse')
		tkinter.ttk.Separator(app, orient=HORIZONTAL).grid(column=0, row=1, columnspan=3, sticky='wen')
		tkinter.ttk.Separator(app, orient=HORIZONTAL).grid(column=0, row=2, columnspan=3, sticky='wen')

				
class ThreadedClient:

    def __init__(self, master):
        self.master = master
        # Create the queue
        self.queue = Queue.Queue()
        # Set up the GUI part
        self.gui = GuiPart(master, self.queue, self.endApplication)
        # Set up the thread to do asynchronous I/O
        # More threads can also be created and used, if necessary
        self.running = 1
        self.thread1 = threading.Thread(target=self.workerThread1)
        self.thread1.start()
        # Start the periodic call in the GUI to check if the queue contains
        # anything
        self.periodicCall()

    def periodicCall(self):
        """
        Check every 200 ms if there is something new in the queue.
        """
        self.gui.processIncoming()
        if not self.running:
            # This is the brutal stop of the system. You may want to do
            # some cleanup before actually shutting it down.
            import sys
            sys.exit(1)
        self.master.after(200, self.periodicCall)

    def workerThread1(self):
        """
        This is where we handle the asynchronous I/O. For example, it may be
        a 'select(  )'. One important thing to remember is that the thread has
        to yield control pretty regularly, by select or otherwise.
        """
        while self.running:
            # To simulate asynchronous I/O, we create a random number at
            # random intervals. Replace the following two lines with the real
            # thing.
            time.sleep(rand.random() * 1.5)
            msg = rand.random()
            self.queue.put(msg)

    def endApplication(self):
        self.running = 0
			
app = Tk()
client = ThreadedClient(app)
rand = random.Random()
app.mainloop()