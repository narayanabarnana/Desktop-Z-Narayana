'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
								'Carepoint NS-LOB (Non Standard) DriverScript
								'Created By : Srirekha Talasila
								'Created On : 12/15/2016
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Dim gstrRelativePath,ActionNumber,ActionCount

Environment.Value("RelativePath") = "\\uszzaschnas01\ClaimsASP\eZaccess\QA\Carepoint\Carepoint_Automation"
gstrRelativePath = Environment.Value("RelativePath")

Environment.Value("ExecutionResultsPath") = gstrRelativePath & "\CP_Results\NS"
Environment.Value("ClaimNumberPath") = Environment.Value("ExecutionResultsPath") 
Environment.Value("RegressionTrackerPath") = Environment.Value("RelativePath") & "\CP_Regression_Tracker\CarePoint _Regression_Tracker.xls"

Environment.Value("Claim_Number") = 0
Environment.Value("SCaseId") = 0

REM ======= Assign values to the variables	================================================

int_TotalScenarioToExecute=0
int_TotalPassedScenarios=0
int_TotalFailedScenarios=0
ActionCount = 16
Environment.Value("str_StepNum") = 0
Environment.Value("counter") = 1

Set objExcel=CreateObject("Excel.Application")

REM ======== Retriving LOB Name from Driver Script =========================================

action_Name = Environment.value("TestName")
LOB_Name = MID(action_Name,1,INSTR(action_Name,"_")-1)

REM ========  Environment File Path ========================================================
Run_Manager_Path = gstrRelativePath &"\CP_EnvFiles\RunManager.xls"

Set objWrkbook=objExcel.Workbooks.Open(Run_Manager_Path)
Set objSheet1=objExcel.Sheets("URL")
Set objSheet2=objExcel.Sheets(LOB_Name)

REM =======  TO GET LOB WISE SHEET NAMES FROM RUN MANAGER ===================================
Dim SheetNames(10)
rowcount2=objSheet2.usedrange.rows.count
sheetcount=0
	For rows=2 To rowcount2
		If objSheet2.Cells(rows,1).Value<>"" Then
			SheetNames(sheetcount)=objSheet2.Cells(rows,1).Value
			sheetcount=sheetcount+1
		End If
	Next

URL_RowCount = objSheet1.usedrange.rows.count

    For Url_Itr = 2 to URL_RowCount
		If objSheet1.Cells(Url_Itr,1).Value = "Y" Then
			App_Identifier = Left(objSheet1.Cells(Url_Itr,3).Value,2)
			If App_Identifier = "CP" Then
				Environment.Value("Current_Region") = objSheet1.Cells(Url_Itr,2).Value
				Environment.Value("CP_URL") = objSheet1.Cells(Url_Itr,4).Value
				Environment.Value("CP_LoginId") = objSheet1.Cells(Url_Itr,5).Value
				Environment.Value("CP_LoginPassword") = objSheet1.Cells(Url_Itr,6).Value
			Else	
				Environment.Value("EZ_URL") = objSheet1.Cells(Url_Itr,4).Value
				Environment.Value("EZ_LoginId") = objSheet1.Cells(Url_Itr,5).Value
				Environment.Value("EZ_LoginPassword") = objSheet1.Cells(Url_Itr,6).Value
			End If
		End If
	Next

objWrkbook.Close
objExcel.Quit
Set objWrkbook=Nothing
Set objSheet1=Nothing
Set objSheet2=Nothing
Set objExcel = Nothing

'Initialize Report
Call Reporting_Initialize()

Dim arrayBusinessObject(30)

'Scenarios XLS Path 
Scenario_Path = gstrRelativePath&"\CP_TestData_RegSuite\"&LOB_Name&"_End-End.xls"
'Import Sheets to Data Table
sheetcount=sheetcount-1
For iterator=0 To sheetcount			
	Set arrayBusinessObject(iterator) = DataTable.AddSheet(SheetNames(iterator))
	DataTable.ImportSheet gstrRelativePath&"\CP_TestData_RegSuite\"&LOB_Name&"_End-End.xls",SheetNames(iterator),SheetNames(iterator)
Next

'TO CHECK FOR RUN FLAG
testCaseCount= arrayBusinessObject(0).GetRowCount() 'DataTable.GetSheet("BusinessFlow").GetRowCount
For counter= 1 to testCaseCount
		Environment.Value("counter") = counter 
		'Set Current Row in Business Flow Sheet
		DataTable.GetSheet("BusinessFlow").SetCurrentRow(counter)
		Environment.Value("caseId") = arrayBusinessObject(0).GetParameter("TestcaseID") ' DataTable.Value("TestcaseID","BusinessFlow")
		strRunFlag = Ucase(Trim(arrayBusinessObject(0).GetParameter("RunFlag")))
		str_MachineName = Ucase(Trim(arrayBusinessObject(0).GetParameter("MachineName")))
		
		'TO CHECK FOR RUN FLAG and EXECUTE THE TEST CASE
		If Ucase(Trim(DataTable.Value("RunFlag","BusinessFlow")))="Y" and UCASE(Environment("LocalHostName")) = UCASE(Trim(DataTable.Value("MachineName","BusinessFlow"))) Then
			ActionNumber = 1
'			SystemUtil.CloseProcessByName "iexplore.exe"
			DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
			Call fn_PreExecution()
			If Instr(DataTable.Value("TestcaseID","BusinessFlow"),"_") > 0 then
				Environment.Value("str_TCID") = DataTable.Value("TestcaseID","BusinessFlow")
				Environment.Value("str_TCName") = Left(DataTable.Value("TestcaseID","BusinessFlow"),Instr(DataTable.Value("TestcaseID","BusinessFlow"),"_")-1)
				Environment.Value("SceNum") = LOB_Name & "_"& Environment("str_TCName")
			End If
			
			Call Reporting_TestCase_Initialize()
		
				 Do While  ActionNumber <= ActionCount  
					 If DataTable.Value("Action_"&ActionNumber,"BusinessFlow")<>"" Then
					 	FunctionName = DataTable.Value("Action_"&ActionNumber,"BusinessFlow")
						Environment.value("FunctionName")=FunctionName
						Print "+++++++++++++++++++++++++++++++++++++++++++++++ " & FunctionName & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
						DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
						Execute FunctionName 
						ActionNumber=ActionNumber+1
					 Else
					 	ActionNumber=ActionNumber+1
					 End If
					
				Loop	
			
			int_TotalScenarioToExecute=int_TotalScenarioToExecute+1	
			
			If Environment.value("FailedSteps")>0 Then
				int_TotalFailedScenarios=int_TotalFailedScenarios+1
			Else
				int_TotalPassedScenarios=int_TotalPassedScenarios+1
			End If
			Print "++++++++++++++++++++++++++++++++++++++++ TestCase END +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ "
			Print "Passed Steps are :  " & Environment.value("PassedSteps")
			Print "Failed Steps are :  " & Environment.value("FailedSteps")
			Environment.Value("int_TotalScenarioToExecute")=int_TotalScenarioToExecute
			Environment.Value("int_TotalPassedScenarios")=int_TotalPassedScenarios
			Environment.Value("int_TotalFailedScenarios")=int_TotalFailedScenarios		
			
			Call fnCloseReport(Date)
			Call fnInsertSection_BatchRun(Environment("HTMLSummaryFile"),Environment("HTMLResultsPath"))
			Call ExcelReport_Generation()
			Call Update_Regression_Tracker()
			Call Clear_Cookies()
		End If
	Next

For count3=0 To sheetcount
	DataTable.DeleteSheet(SheetNames(count3))
Next

Print NOW


'HTML_Execution_Summary_Close
Call fnCloseReport_BatchRun(Environment.Value("HTMLSummaryFile"))




