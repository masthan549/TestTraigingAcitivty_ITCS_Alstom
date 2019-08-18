from tkinter import Label, Button, Entry
import tkinter as tk
from tkinter import messagebox, filedialog, PhotoImage, StringVar, SUNKEN, W, X, BOTTOM
from os import path
import TestTriagingAnalysis
import threading, sys
import os
from tkinter.filedialog import askopenfilename

class GUI_COntroller:
    '''
	   This class initialize the required controls for TkInter GUI
	'''
    def __init__(self,TkObject):
 
 
	    #Load company image
        Imageloc=tk.PhotoImage(file='alstom_logo.gif')		
        label3=Label(image=Imageloc,)
        label3.image = Imageloc		
        label3.place(x=200,y=10)
		

        global TkObject_ref, validPrevBuildDirSelected, selectedDir_seq, validtriagingDirSelected, validExcelSelected
        TkObject_ref =  TkObject
        validtriagingDirSelected = False		
        validPrevBuildDirSelected = False
        validExcelSelected = False		

		
        #label
        global label1		
        label1 = Label(TkObject,bd=7, text="Test triaging report preparation for triage folder and previous results", bg="green", fg="black",width=60, font=200)	
        label1.place(x=100,y=80)
        label1.config(font=('helvetica',12,'bold'))		

        global label_select_triage_loc		
        label_select_triage_loc = Label(TkObject,bd=7, text="1. Select triage HTML files directory :: ", bg="orange", fg="black",width=30, font=200)	
        label_select_triage_loc.place(x=30,y=140)
        label_select_triage_loc.config(font=('helvetica',12,'bold'))			
		
        #select sequence files directory
        global 	button1_select_triage_loc	
        button1_select_triage_loc=Button(TkObject,activebackground='green',borderwidth=5, text='Click here to select path',width=25, command=GUI_COntroller.selectSeqDirectory)
        button1_select_triage_loc.place(x=430,y=140)
        button1_select_triage_loc.config(font=('helvetica',12,'bold'))

        global label_select_prevBuild_loc		
        label_select_prevBuild_loc = Label(TkObject,bd=7, text="2. Select previous results directory :: ", bg="orange", fg="black",width=30, font=200)	
        label_select_prevBuild_loc.place(x=30,y=260)
        label_select_prevBuild_loc.config(font=('helvetica',12,'bold'))			
		
        #select sequence files directory
        global 	button1_select_prevBuild_loc	
        button1_select_prevBuild_loc=Button(TkObject,activebackground='green',borderwidth=5, text='Click here to select path',width=25, command=GUI_COntroller.selectResDirectory)
        button1_select_prevBuild_loc.place(x=430,y=260)
        button1_select_prevBuild_loc.config(font=('helvetica',12,'bold'))		
		
		
        #label
        global label_selectFile		
        label_selectFile = Label(TkObject,bd=7, text="3. Select the sequence list excel file ::", bg="orange", fg="black",width=30, font=200)	
        label_selectFile.place(x=30,y=380)
        label_selectFile.config(font=('helvetica',12,'bold'))			

        #select file
        global 	button_selectFile
        button_selectFile=Button(TkObject,activebackground='green',borderwidth=5, text='Select file!!',width=25, command=GUI_COntroller.openfile)
        button_selectFile.place(x=430,y=380)
        button_selectFile.config(font=('helvetica',12,'bold'))		
		

        #Exit Window
        global button2_close		
        button2_close=Button(TkObject,activebackground='green',borderwidth=5, text='Close Window', command=GUI_COntroller.exitWindow)
        button2_close.place(x=600,y=510)	
        button2_close.config(font=('helvetica',12,'bold'))	

				
        #select sequence files directory
        global 	button1_executeTest	
        button1_executeTest=Button(TkObject,activebackground='green',borderwidth=5, text='Run Test triage Analyse',width=25, command=GUI_COntroller.RunTest)
        button1_executeTest.place(x=230,y=510)
        button1_executeTest.config(font=('helvetica',12,'bold'))	

    def exitWindow():
            TkObject_ref.destroy()

    def RunTest():

        runTest = True
        global validtriagingDirSelected, validPrevBuildDirSelected, validExcelSelected
	
        if validtriagingDirSelected == False:
            messagebox.showerror('Error','Please select a test triage HTML files directory!')
            runTest = False
	
        elif validPrevBuildDirSelected == False:
            messagebox.showerror('Error','Please select a previous build test results directory!')
            runTest = False			

        elif validExcelSelected == False:			
            messagebox.showerror('Error','Select valid xls/xlsx file!!')			
            runTest = False				
			
        if (runTest and (GUI_COntroller.checkHTMLFilesExist() == True)):
            TestTriaging.RunTest()
	
			
    def checkHTMLFilesExist():

        HTMLFilesExist = True
        testReportsPath_currentBuild = selectedDir_seq+"\\"
        testReportsPath_prevBuild = selectedDir_res+"\\"
	    
        lstFile = sorted([file for file in os.listdir(testReportsPath_currentBuild) if file.endswith('.html')])
        totalNumberOfReportsCurrentBuild = len(lstFile)
        reportNamesFromPrevBuild = sorted([file for file in os.listdir(testReportsPath_prevBuild) if file.endswith('.html')])
	    
	    
        if (totalNumberOfReportsCurrentBuild == 0):
            messagebox.showerror('Error','HTML reports NOT available in selected current build ditectory, please try again!!')
            HTMLFilesExist = False			
        elif (len(reportNamesFromPrevBuild) == 0):
            messagebox.showerror('Error','HTML reports NOT available in selected previous build ditectory, please try again!!') 	
            HTMLFilesExist = False
        
        return HTMLFilesExist		
			
    def selectSeqDirectory():
            global selectedDir_seq, validtriagingDirSelected
            currdir = os.getcwd()
            selectedDir_seq = filedialog.askdirectory(initialdir=currdir, title='Please select a directory')
            if not path.isdir(selectedDir_seq):
                messagebox.showerror('Error','Please select a valid directory!')				
            else:

                label4= Label(TkObject_ref,bg='white',text=str(selectedDir_seq),font=40)
                label4.place(x=80,y=200)
                validtriagingDirSelected = True

    def selectResDirectory():
            global selectedDir_res, validPrevBuildDirSelected
            currdir = os.getcwd()
            selectedDir_res = filedialog.askdirectory(initialdir=currdir, title='Please select a directory')			
            if not path.isdir(selectedDir_res):
                messagebox.showerror('Error','Please select a valid directory!')				
            else:

                label5= Label(TkObject_ref,bg='white',text=str(selectedDir_res),font=40)
                label5.place(x=80,y=320)
                validPrevBuildDirSelected = True	
				

    def openfile():
        global filepath,filepath_temp, validExcelSelected	
        filepath = askopenfilename()	
        filepath_temp=filepath.split('/')
        filepath_temp=filepath_temp[len(filepath_temp)-1]
        validExcelSelected = True
		
        if not (filepath_temp.endswith('xls') or filepath_temp.endswith('xlsx')):
            validExcelSelected = False
        else:
            label6= Label(TkObject_ref,bg='white',text=str(filepath),font=40)
            label6.place(x=80,y=440)
            validExcelSelected = True
			
class TestTriaging:
    def RunTest(): 

        global thread,statusBarText, button1_executeTest

        button1_executeTest.config(state="disabled")
		
        statusBarText = StringVar()		
        StatusLabel = Label(TkObject_ref, textvariable=statusBarText, fg="green", bd=1,relief=SUNKEN,anchor=W) 
        StatusLabel.config(font=('helvetica',11,'bold'))
        StatusLabel.pack(side=BOTTOM, fill=X)
        statusBarText.set("Test traiging report analysis in progress... please wait...")
		
        thread = threading.Thread(target=TestTriagingAnalysis.script_exe, args = (selectedDir_seq,selectedDir_res,filepath, TkObject_ref,statusBarText))
        thread.start()

if __name__ == '__main__':	
	
       root = tk.Tk()
       
       #Change the background window color
       root.configure(background='gray')     
       
       #Set window parameters
       root.geometry('850x680')
       root.title('Test triaging activity')
       
       #Removes the maximizing option
       root.resizable(0,0)
       
       ObjController = GUI_COntroller(root)
       
       #keep the main window is running
       root.mainloop()
       sys.exit()
