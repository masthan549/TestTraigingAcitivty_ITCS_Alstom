import sys, os, re, xlsxwriter
from datetime import datetime 
from tkinter import messagebox
from xlrd import open_workbook

worksheet_html = ""
worksheet_html_start_rowCounter = 0
workbook = ""
testSheetName = ""
SeqExecutionStatus_Final = {}
SeqNames_Final = []
skippedTestList = []

def getIndividualTestCaseStatus(completeTestReportPath):

    # This list only contains failed test case number and its test steps
    FailedTestSteps = []
    
    if os.path.isfile(completeTestReportPath) and completeTestReportPath.endswith(".html"):

        #Fetch the line numbers from HTML file where test failed    
        rePattern = re.compile('<FONT SIZE="-1">Failed</FONT>')
        FileLines = open(completeTestReportPath)
        lineNumberList = []
        FailedTestSteps_temp = []
        
        for (LineIndex, LineText) in enumerate(FileLines):
           if rePattern.match(LineText):
                lineNumberList.append(LineIndex-13)           

        FileLines.close()

        #Fetch only test steps which are failed         
        with open(completeTestReportPath) as fp:
            for indx, line in enumerate(fp):
                if (indx+1) in lineNumberList:
                    FailedTestSteps_temp.append(line.strip())                    

        #Remove unwanted data from failed steps from HTML
        FailedTestSteps_temp    = [listItem.replace('<TD valign="top" COLSPAN="2" BGCOLOR="#00C4C4">','') for listItem in FailedTestSteps_temp]
        FailedTestSteps_temp    = [listItem.replace('</TD>','') for listItem in FailedTestSteps_temp]

        for listItem in FailedTestSteps_temp:
            reEx = re.search(r'>(.*)</A>',listItem,re.M|re.IGNORECASE)
            try:
                FailedTestSteps.append((reEx.groups()[0]).strip())
            except:FailedTestSteps.append(listItem.strip())
        
    return FailedTestSteps     


def seeTheNumberOfrepitionsOfReport(reportName, allReportsList):

    
    sameReportNames = []
    reportNameAndIndx = []
    reportCounter = -1
    numberOfTermResults = 0
    
    #See how many reports available for the same test
    testSeqName = reportName.split("_Report")[0]
    reportNameAndIndx.append(reportName)
    reportNameAndIndx.append(reportCounter)    
    sameReportNames.append(reportNameAndIndx)      
    
    for lineIndx, testReportName_comp in enumerate(allReportsList):
        reportNameAndIndx = []    

        if ((testSeqName+"_Report") in testReportName_comp) and (reportName != testReportName_comp):
            reportNameAndIndx.append(testReportName_comp)
            reportNameAndIndx.append(lineIndx)    
            sameReportNames.append(reportNameAndIndx)		
    return sameReportNames
            
def fetchMoreNumberOfFailuresReportFromMultipleReports(testReportPath_arg, listRepitition_arg):

    selectedFailedReports = [] 
    resultsOfReport = []
    testResultsConsistent = True	

    if len(listRepitition_arg) > 0:
        testReportPathAndName = testReportPath_arg+listRepitition_arg[0][0]
        failedTestSteps_Prev = getIndividualTestCaseStatus(testReportPathAndName)    
        selectedFailedReports = failedTestSteps_Prev    
        resultsOfReport = listRepitition_arg[0][0]
        
        for lstIndx in listRepitition_arg[1:]:
            testReportPathAndName = testReportPath_arg+lstIndx[0]
            failedTestSteps_Next = getIndividualTestCaseStatus(testReportPathAndName)
            
            #Two or more than two test reports didnt match so test results marked as not consistent. But tool pics up the report which has more number of failures
            if((len(failedTestSteps_Next) > len(failedTestSteps_Prev))):
                selectedFailedReports = failedTestSteps_Next
                resultsOfReport = lstIndx[0]
                testResultsConsistent = False
        
            # This condition will be set to TRUE when two reports have same number of failures. But diff steps have the failures
            if(set(failedTestSteps_Next) != set(failedTestSteps_Prev)): 
                testResultsConsistent = False
            
    return (testResultsConsistent, selectedFailedReports, resultsOfReport)
            
def findTestResultsStatus(reportName):

    ResultsStatus = "NA"

    if(reportName.find("[P].html")) > -1:
        ResultsStatus = "PASS"

    if(reportName.find("[F].html")) > -1:
        ResultsStatus = "FAIL"

    if(reportName.find("[T].html")) > -1:
        ResultsStatus = "TERMINATED"
            
    return ResultsStatus
	
	
def writeExcelHeader():

	
    global worksheet_html
    global worksheet_html_start_rowCounter
    global workbook, testSheetName
	
    dt = datetime.now()
    day_hour_min_sec = str(dt.day)+"_"+str(dt.hour)+"_"+str(dt.minute)+"_"+str(dt.second)
    testSheetName = 'TestResultSheet'+'_'+str(day_hour_min_sec)+'.xlsx'	
	
    # Creating Xlsheet to write the data into that if it is not exist 	
    try:
        workbook = xlsxwriter.Workbook('TestResultSheet'+'_'+str(day_hour_min_sec)+'.xlsx')
    except:
        messagebox.showerror('Error','Please close the Workbook and try again!') 
        TkObject.destroy()
        sys.exit()
    
    # Colors	
    bold = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    merge_format1 = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'size': 14, 'bg_color': '#d8d6d2', 'text_wrap': True})	
    merge_format2 = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'size': 14, 'bg_color': '#55aadd', 'text_wrap': True})	
    merge_format3 = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'size': 14, 'bg_color': 'purple', 'text_wrap': True})	
    merge_format4 = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'size': 14, 'bg_color': '#f56329', 'text_wrap': True})	

   
    # Write the sequence and testcase and its result into workbook sheet
    worksheet_html = workbook.add_worksheet('TestResultsComparision')
    worksheet_html_start_rowCounter = 14
	
    worksheet_html.merge_range('B12:G12', 'Current build Test Reports from triage fodler', merge_format1)		
    worksheet_html.write('H12',"Previous build Test Results", merge_format2)
    worksheet_html.merge_range('I12:I13', 'Test Results SAME in previous and current build (triage folder)?', merge_format3)		
    worksheet_html.merge_range('J12:K12', 'Failed test steps comparision between previous build and current build (traige folder)', merge_format4)		
	
    worksheet_html.write('B13',"Test Sequence", bold)
    worksheet_html.write('C13',"Test Report (Considered) from Triage", bold)
    worksheet_html.write('D13',"Test Result", bold)
    worksheet_html.write('E13',"Number of Runs got", bold)
    worksheet_html.write('F13',"Test Results are consistent in triage folder from current build?", bold)
    worksheet_html.write('G13',"Number of PASS/FAIL/TERMINATED Results in triage folder", bold)
    worksheet_html.write('H13',"Test Result", bold)
    worksheet_html.write('J13',"Test Steps Failed in previous build", bold)
    worksheet_html.write('K13',"Test Steps Failed in Current build (in triage folder)", bold)

    # Set width of the columns
    worksheet_html.set_column(0,0,6)	
    worksheet_html.set_column(1,1,20)	
    worksheet_html.set_column(2,2,20)	
    worksheet_html.set_column(3,3,14)	
    worksheet_html.set_column(4,4,10)	
    worksheet_html.set_column(5,5,20)	
    worksheet_html.set_column(6,6,18)	
    worksheet_html.set_column(7,7,12)	
    worksheet_html.set_column(8,8,16)	
    worksheet_html.set_column(9,9,20)	
    worksheet_html.set_column(10,10,20)	
    worksheet_html.set_row(12,73)	
    worksheet_html.set_row(11,81)	

	
def writeResultsIntoReport(testSeqName_to_report, selectedReportName_to_report, selectedReportStatus_to_report, listRepitition_to_report, resultsConsistent_to_report, testResultsNumber_to_report, prevResultsStatus_to_report, resMatches, prevBuildFailedSteps, currBuildFailedSteps):

	
    global worksheet_html
    global worksheet_html_start_rowCounter
    global workbook
	
    bold_green_format = workbook.add_format({'bg_color': '#15FD0D', 'bold': 1})	
    bold_red_format = workbook.add_format({'bg_color': 'red', 'bold': 1})	
    bold_orange_format = workbook.add_format({'bg_color': '#feba29', 'bold': 1})	
    bold_yellow_format = workbook.add_format({'bg_color': 'yellow', 'bold': 1})	
    bold_grey_format = workbook.add_format({'bg_color': '#D7DBDD', 'bold': 1})	
    bold_blue_format = workbook.add_format({'bg_color': '#AED6F1', 'bold': 1})	
    
    worksheet_html.write('B'+str(worksheet_html_start_rowCounter),testSeqName_to_report)
    worksheet_html.write('C'+str(worksheet_html_start_rowCounter),selectedReportName_to_report)
    worksheet_html.write('D'+str(worksheet_html_start_rowCounter),selectedReportStatus_to_report)
    worksheet_html.write('E'+str(worksheet_html_start_rowCounter),listRepitition_to_report)
	
    #Color code the test results status
    if(resultsConsistent_to_report == False):	
        worksheet_html.write('F'+str(worksheet_html_start_rowCounter),resultsConsistent_to_report, bold_red_format)
    else:
        worksheet_html.write('F'+str(worksheet_html_start_rowCounter),resultsConsistent_to_report)

    worksheet_html.write('G'+str(worksheet_html_start_rowCounter),testResultsNumber_to_report)
    worksheet_html.write('H'+str(worksheet_html_start_rowCounter),prevResultsStatus_to_report)
	
    #Color code the test results status
    if(resMatches == "SKIP"):
        worksheet_html.write('I'+str(worksheet_html_start_rowCounter),resMatches, bold_grey_format)   
    elif(resMatches == "NO RESULTS"):
        worksheet_html.write('I'+str(worksheet_html_start_rowCounter),"NO RESULTS IN TRIAGE", bold_blue_format) 		
    elif(resMatches == "PASS"):
        worksheet_html.write('I'+str(worksheet_html_start_rowCounter),resMatches, bold_green_format)
    elif(resMatches == "FAIL"):
        worksheet_html.write('I'+str(worksheet_html_start_rowCounter),resMatches, bold_red_format)		
    elif(resMatches == "C-PASS"):
        worksheet_html.write('I'+str(worksheet_html_start_rowCounter),resMatches, bold_orange_format)		
    elif(resMatches == "PREVIOUS BUILD HAS NO REPORT FOR THIS TEST"):
        worksheet_html.write('I'+str(worksheet_html_start_rowCounter),resMatches, bold_yellow_format)
		
    worksheet_html.write('J'+str(worksheet_html_start_rowCounter),prevBuildFailedSteps)
    worksheet_html.write('K'+str(worksheet_html_start_rowCounter),currBuildFailedSteps)
    worksheet_html.write('L'+str(worksheet_html_start_rowCounter)," ") # This is just to avoid overlapping the text over next cells
    worksheet_html_start_rowCounter = worksheet_html_start_rowCounter + 1	

def writeMetricsIntoSheet():
			
    global worksheet_html
    global workbook	
    global SeqNames_Final, worksheet_html_start_rowCounter
	
    bold_format = workbook.add_format({'bold': 1})
    bold_green_format = workbook.add_format({'bg_color': '#15FD0D', 'bold': 1})	
    bold_green_format = workbook.add_format({'bg_color': '#15FD0D', 'bold': 1, 'num_format': '0.00%'})	
    bold_red_format = workbook.add_format({'bg_color': 'red', 'bold': 1})
    bold_yellow_format = workbook.add_format({'bg_color': 'yellow', 'bold': 1})	
    bold_grey_format = workbook.add_format({'bg_color': '#D7DBDD', 'bold': 1})		
    bold_grey_format = workbook.add_format({'bg_color': '#D7DBDD', 'bold': 1, 'num_format': '0.00%'})		
	
    #Total tests
    worksheet_html.write('B2',len(SeqNames_Final), bold_format)	
    worksheet_html.merge_range('C2:F2',"Tests to Be executed in Sweep", bold_format)
	
    #test status
    worksheet_html.write('B3',"=COUNTIF($I$14:$I$"+str(worksheet_html_start_rowCounter-1)+",\"PASS\")")	
    worksheet_html.merge_range('C3:F3',"Pass")    
	
    worksheet_html.write('B4',"=COUNTIF($I$14:$I$"+str(worksheet_html_start_rowCounter-1)+",\"C-PASS\")")	
    worksheet_html.merge_range('C4:F4',"C-Pass") 
	
    worksheet_html.write('B5',"=SUM(B3:B4)")	
    worksheet_html.merge_range('C5:F5',"Credit Taken For…")	
	
    worksheet_html.write('B6',"=SUM(COUNTIF($I$14:$I$"+str(worksheet_html_start_rowCounter-1)+",\"FAIL\"),COUNTIF($I$14:$I$"+str(worksheet_html_start_rowCounter-1)+",\"PREVIOUS BUILD HAS NO REPORT FOR THIS TEST\"),COUNTIF($I$14:$I$"+str(worksheet_html_start_rowCounter-1)+",\"NO RESULTS\"))")	
    worksheet_html.merge_range('C6:F6',"Needs Analysis (Investigations)")		

    worksheet_html.write('B7',"=(B3+B9)/B2", bold_green_format)	
    worksheet_html.merge_range('C7:F7',"% Complete With Passing Results", bold_green_format)

    worksheet_html.write('B8',"=(B3+B4+B9+B6)/B2", bold_grey_format)	
    worksheet_html.merge_range('C8:F8',"% Complete either Passing/C-Pass/Fail or Investigation", bold_grey_format)	
	
    worksheet_html.write('B9',"=COUNTIF($I$14:$I$"+str(worksheet_html_start_rowCounter-1)+",\"FAIL\")", bold_red_format)	
    worksheet_html.merge_range('C9:F9',"Tests Failure tied to SW CRs", bold_red_format)
	
    worksheet_html.write('B10',"=COUNTIF($I$14:$I$"+str(worksheet_html_start_rowCounter-1)+",\"SKIP\")", bold_yellow_format)
    worksheet_html.merge_range('C10:F10',"Skipped Tests", bold_yellow_format) 	
	
def getReportNamesWithoutTerm(listRepitition_rec):

    termResultIndxPos = []
    resWithoutTerm = []		
    listWithoutTermResults = list(listRepitition_rec)	
	
    for indx in range(0,len(listWithoutTermResults)):
        if((listWithoutTermResults[indx][0].find('[T].html')) > -1):
            termResultIndxPos.append(indx)              

    for indx in range(0,len(listWithoutTermResults)):
        if indx not in termResultIndxPos:
            resWithoutTerm.append(listWithoutTermResults[indx])
			
    return resWithoutTerm
	

def getSeqNamesListFromSheet(reportPath):
#if __name__ == "__main__":

    #reportPath = "C:\\Users\\402096\\My_Data\\Work\\ITCS\\Automation\\Test Results Moving from triage\\"
    #workbookName = "SeqNameAndStatus.xlsx"
	
    xlsPtr = open_workbook(reportPath)
    HLTP_Template = xlsPtr.sheet_by_index(0)
    numberOfRows = HLTP_Template.nrows
    numberOfCols = HLTP_Template.ncols
    global SeqExecutionStatus_Final
    global SeqNames_Final
    global skippedTestList

    for rowIndx in range(0,numberOfRows):
        cellValueSeq = (HLTP_Template.cell(rowIndx,0).value).strip()
        cellValueStatus = (HLTP_Template.cell(rowIndx,1).value).strip()
        SeqNames_Final.append(cellValueSeq.upper())
        if(cellValueStatus == ""):
            SeqExecutionStatus_Final[cellValueSeq.upper()] = ["","","","","","","NO RESULTS","",""]
        else:
            SeqExecutionStatus_Final[cellValueSeq.upper()] = ["","","","","","","SKIP","",""]
            skippedTestList.append(cellValueSeq.upper())
		

def script_exe(testReportsPath_currentBuild_arg, testReportsPath_prevBuild_arg, filepath, TkObject_ref, statusBarText):

    global TkObject,statusBar, SeqExecutionStatus_Final, SeqNames_Final, skippedTestList
	
    TkObject = TkObject_ref
    statusBar = statusBarText
    testReportsPath_currentBuild = testReportsPath_currentBuild_arg+"\\"
    testReportsPath_prevBuild = testReportsPath_prevBuild_arg+"\\"
	

    #testReportsPath_currentBuild = "C:\\Users\\402096\\My_Data\\Work\\ITCS\\Automation\\Test Results Moving from triage\\Results_OE\\Results_acceptableruns\\"    
    #testReportsPath_prevBuild = "C:\\Users\\402096\\My_Data\\Work\\ITCS\\Automation\\Test Results Moving from triage\\Results_0D\\Results_acceptableruns\\"
    
    lstFile = sorted([file for file in os.listdir(testReportsPath_currentBuild) if file.endswith('.html')])
    progressCounter = 0
    totalNumberOfReportsCurrentBuild = len(lstFile)
    reportNamesFromPrevBuild = sorted([file for file in os.listdir(testReportsPath_prevBuild) if file.endswith('.html')])
	
    #Read sequence names and its status
    getSeqNamesListFromSheet(filepath)
	
    if (totalNumberOfReportsCurrentBuild == 0):
        messagebox.showerror('Error','HTML reports NOT available in selected current build ditectory, please try again!!') 
        TkObject.destroy()	
        workbook.close()		
        sys.exit()

    if (len(reportNamesFromPrevBuild) == 0):
        messagebox.showerror('Error','HTML reports NOT available in selected previous build ditectory, please try again!!') 
        TkObject.destroy()			
        workbook.close()		
        sys.exit()
	
    #Write Excel header
    writeExcelHeader()
		
    for testReportName in lstFile:
	
        ############################### CURRENT BUILD TEST ANALYSIS ########################################
		
        #Fetch test sequence name
        testSeqName = testReportName.split("_Report")[0]  
        listWithoutTermResults = []	
        numberOfTermResults = 0		

        #Analyse tests only if it is not SKIP
        if testSeqName.upper() not in skippedTestList:
		
            #Check for Report repitions
            listRepitition = seeTheNumberOfrepitionsOfReport(testReportName, lstFile)
		    
            #List without terminated results
            listRepitition_withoutTerm = getReportNamesWithoutTerm(listRepitition)		
		    
            #Fetch the report which has more number of failures from multiple reports of same test
            resultsConsistent, FailedReportResults, selectedReportName = fetchMoreNumberOfFailuresReportFromMultipleReports(testReportsPath_currentBuild, listRepitition_withoutTerm)
            
            numberOfPASSResults = 0
            numberOfFAILResults = 0
            numberOfTerminatedResults = 0    
            
            #Find out number of failed, PASS and Teminated results we have for every test        
            for reportNames in listRepitition_withoutTerm:
                if (reportNames[0].find("[P].html") > -1):
                    numberOfPASSResults = numberOfPASSResults+1
            
                if (reportNames[0].find("[F].html") > -1):
                    numberOfFAILResults = numberOfFAILResults+1
            
            numberOfTerminatedResults = len(listRepitition) - len(listRepitition_withoutTerm)
            
            testResultsNumber = "Number Of PASS Results: "+str(numberOfPASSResults)+"\nNumber Of FAIL Results: "+str(numberOfFAILResults)+"\nNumber Of Terminated Results: "+str(numberOfTerminatedResults)
                    
            #Remove Duplicate names from main list. Reason for removal of this name is above function pics up the report file which has more number failures.
            for indx_remove in listRepitition[1:]:
                progressCounter = progressCounter + 1            
                lstFile.remove(indx_remove[0])
            
            statusBar.set("Number of HTML files Analysed so far: ("+str(progressCounter)+"/"+str(totalNumberOfReportsCurrentBuild)+")."+"\nCurrently Analysing sequence: \""+ testReportName+"\". Please wait...")    
            progressCounter = progressCounter+1 
		    
            ############################### PREVIOUS BUILD TEST REPORT ANALYSIS ########################################
            
            prevBuildReportName = []        
            
            #Fetch all previous build report names which matches with sequence name 
            for indx_prevBuild in reportNamesFromPrevBuild:
                if testSeqName+"_Report" in indx_prevBuild:
                    prevBuildReportName.append(indx_prevBuild)
                    
            #Check whether results were PASS or FAIL in previous build. If previous build has multiple reports with PASS and FAIL then throw error
            numberOfPASSResultsPrevBuild = 0
            numberOfFAILResultsPrevBuild = 0
            numberOfTerminatedResultsPrevBuild = 0
            prevResultsStatus = "NA"	
		    
            for indxPrevbuildIndx in prevBuildReportName:
                if (indxPrevbuildIndx.find("[P].html") > -1):
                    numberOfPASSResultsPrevBuild = numberOfPASSResultsPrevBuild+1
            
                if (indxPrevbuildIndx.find("[F].html") > -1):
                    numberOfFAILResultsPrevBuild = numberOfFAILResultsPrevBuild+1
            
                if (indxPrevbuildIndx.find("[T].html") > -1):
                    numberOfTerminatedResultsPrevBuild = numberOfTerminatedResultsPrevBuild+1
            
            if (len(prevBuildReportName) == 0):
                prevResultsStatus = "RESULTS NOT AVAILABLE"
                
            elif  not (((numberOfPASSResultsPrevBuild > 0) and (numberOfFAILResultsPrevBuild == 0) and (numberOfTerminatedResultsPrevBuild == 0)) or ((numberOfPASSResultsPrevBuild == 0) and (numberOfFAILResultsPrevBuild > 0) and (numberOfTerminatedResultsPrevBuild == 0)) or ((numberOfPASSResultsPrevBuild == 0) and (numberOfFAILResultsPrevBuild == 0) and (numberOfTerminatedResultsPrevBuild > 0))):
            
                messagebox.showerror('Error',"Multiple Test reports available in previous build for test : "+str(testSeqName)+", Please have only one required report and delete other reports for this test.") 
                TkObject.destroy()
                workbook.close()			
                sys.exit()
                
            #Check for results only if reports available
            if (len(prevBuildReportName) > 0):
                failedTestSteps_PrevBuild = getIndividualTestCaseStatus(testReportsPath_prevBuild+prevBuildReportName[0])
                prevResultsStatus = findTestResultsStatus(prevBuildReportName[0])
            else:
                failedTestSteps_PrevBuild = []
                prevResultsStatus = "RESULTS NOT AVAILABLE"  
            
            ############################### FINAL TEST RESULTS into REPORT #######################################
            
            #If test has only terminated results
            if len(selectedReportName) > 0 :
            
                #Current Build information
                testSeqName_to_report = testSeqName
                selectedReportName_to_report = selectedReportName
                selectedReportStatus_to_report = findTestResultsStatus(selectedReportName)
                listRepitition_to_report = (len(listRepitition)- numberOfTerminatedResults)
                resultsConsistent_to_report = resultsConsistent
                testResultsNumber_to_report = testResultsNumber
                
                #Current build info
                prevResultsStatus_to_report = prevResultsStatus
                
                #Test Results match between previous and current build
                resMatches = "SAME" if ((set(FailedReportResults) == set(failedTestSteps_PrevBuild)) and (len(FailedReportResults) == len(failedTestSteps_PrevBuild))) else "FAIL"
                resMatches = "PREVIOUS BUILD HAS NO REPORT FOR THIS TEST" if (prevResultsStatus == "RESULTS NOT AVAILABLE") else resMatches
                
                #Failed/terminated test steps from previous build
                prevBuildFailedSteps = " "		
                for indxSteps in failedTestSteps_PrevBuild:
                    prevBuildFailedSteps = prevBuildFailedSteps+indxSteps+"\n"
                
                
                #Failed/terminated test steps from current build
                currBuildFailedSteps = " "
                for indxSteps in FailedReportResults:
                    currBuildFailedSteps = currBuildFailedSteps+indxSteps+"\n"      
            
                #Check for C-PASS
                if ((prevResultsStatus_to_report == "FAIL") and (selectedReportStatus_to_report == "FAIL") and (resMatches == "SAME")):
                    resMatches = "C-PASS"
                elif ((prevResultsStatus_to_report == "PASS") and (selectedReportStatus_to_report == "PASS") and (resMatches == "SAME")):
                    resMatches = "PASS"
                
                #writeResultsIntoReport(testSeqName_to_report, selectedReportName_to_report, selectedReportStatus_to_report, listRepitition_to_report, resultsConsistent_to_report, testResultsNumber_to_report, prevResultsStatus_to_report, resMatches, prevBuildFailedSteps, currBuildFailedSteps)
                SeqExecutionStatus_Final[testSeqName_to_report.upper()]=[selectedReportName_to_report, selectedReportStatus_to_report, listRepitition_to_report, resultsConsistent_to_report, testResultsNumber_to_report, prevResultsStatus_to_report, resMatches, prevBuildFailedSteps, currBuildFailedSteps]
            
            else:
            
                #Failed/terminated test steps from previous build
                prevBuildFailedSteps = " "
                for indxSteps in failedTestSteps_PrevBuild:
                    prevBuildFailedSteps = prevBuildFailedSteps+indxSteps+"\n"		
		    
                #writeResultsIntoReport(testSeqName, "THIS TEST HAS ONLY TERMINATED RESULTS", "TERMINATED", len(listRepitition_withoutTerm), True, testResultsNumber, prevResultsStatus, "FAIL", prevBuildFailedSteps, "")
            
                SeqExecutionStatus_Final[testSeqName.upper()] = ["THIS TEST HAS ONLY TERMINATED RESULTS", "TERMINATED", len(listRepitition_withoutTerm), True, testResultsNumber, prevResultsStatus, "FAIL", prevBuildFailedSteps, ""]
            
			
    #Write results into sheet
    for seqName in SeqNames_Final:
        key = seqName
        value = SeqExecutionStatus_Final[key]
        writeResultsIntoReport(key,value[0],value[1],value[2],value[3],value[4],value[5],value[6],value[7],value[8])
		
    #Write metrics into sheet
    writeMetricsIntoSheet()
			
    # add borders to sheet
    border_format=workbook.add_format({'border':1})	
    worksheet_html.conditional_format('B12:K'+str(worksheet_html_start_rowCounter-1),{ 'type' : 'no_blanks' , 'format' : border_format})	
    worksheet_html.conditional_format('B12:K'+str(worksheet_html_start_rowCounter-1),{ 'type' : 'blanks' , 'format' : border_format})	
    worksheet_html.conditional_format('B2:F10',{ 'type' : 'no_blanks' , 'format' : border_format})	
	
    workbook.close()

    statusBarText.set("DONE!!")	
    messagebox.showinfo('DONE!!',"Test triaging analysis sheet produced in : "+str(os.getcwd())+"\\"+testSheetName)
    TkObject.destroy()
    sys.exit()