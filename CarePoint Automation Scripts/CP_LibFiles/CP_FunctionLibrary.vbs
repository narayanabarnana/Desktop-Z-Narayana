'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
								'Carepoint Register User Functions
								'Created By : Srirekha Talasila
								'Created On : 12/05/2016
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Rem  === Register User Functions
Rem*******************register for setting or selecting value***********************
RegisterUserFunc "WebEdit", "Set", "gfReg_SetorSelect"
RegisterUserFunc "WebCheckBox","Set","gfReg_SetorSelect"
RegisterUserFunc "WebRadioGroup","Select","gfReg_SetorSelect"
RegisterUserFunc "WebList","Select","gfReg_SetorSelect"

Rem*******************register for Clicking Element***********************
RegisterUserFunc "WebButton","Click","gfReg_Click"
RegisterUserFunc "Link","Click","gfReg_Click"
RegisterUserFunc "WebElement","Click","gfReg_Click"
RegisterUserFunc "WebRadioGroup","Click","gfReg_Click"
RegisterUserFunc "WebCheckBox","Click","gfReg_Click"
RegisterUserFunc "Image","Click","gfReg_Click"

Public Function fn_PreExecution()
		Print NOW
		Environment.Value("NewClaimNumber")=0
		Environment.value("TotalSteps")=0
		Environment.value("TotalExecutedSteps")=0
		Environment.value("StepsRemaining")=0
		Environment.value("PassedSteps")=0
		Environment.value("FailedSteps")=0
		Environment.value("str_ExecutionTime")=0
End Function
		
'################################################################################################################
Rem===FunctionName - fn_UpdateTestResults
Rem ===Description:- This Function is to update the Results
Rem ===Designed By:- Srirekha Talasila
'################################################################################################################
Public Function fn_UpdateTestResults(str_ScreenName,str_Operation,str_Status,ErrDescription)
	Err.Clear  
'	On Error Resume Next
	Print int_StepNum & chr(32) & str_ScreenName & chr(32) & str_Operation & chr(32) & str_Status & chr(32) & chr(32) & ErrDescription
	Environment.Value("str_StepNum")=Environment.Value("str_StepNum")+1
	str_SnapshotsPath=Environment.Value("ScreenShotPath") & "\Step_" & Environment.Value("str_StepNum") & ".png"
	Environment.Value("SnapshotsPath")=str_SnapshotsPath
	wait(1)
   	Desktop.CaptureBitmap str_SnapshotsPath,TRUE  
	Environment.Value("int_StepNum")=Environment.Value("str_StepNum")
	Environment.Value("str_SheetName")=str_ScreenName
	Environment.Value("str_Operation")=str_Operation
	Environment.Value("str_Status")=str_Status
	Environment.Value("ErrDescription")=ErrDescription
	Call fnInsertSection()
	Environment.value("TotalExecutedSteps")=Environment.value("TotalExecutedSteps")+1
	If str_Status="FAIL" Then
		Environment.value("FailedSteps")=Environment.value("FailedSteps")+1
		Print "FailedCount:" & chr(32) & Environment.value("FailedSteps")
	ElseIf str_Status="PASS" Then
		Environment.value("PassedSteps")=Environment.value("PassedSteps")+1
		Print "PassedCount:" & chr(32) & Environment.value("PassedSteps")
	End If
End Function

'################################################################################################################
Rem===FunctionName - fn_ObjectExist
Rem ===Description:- This Function for Verifying object Exist or not
Rem ===Designed By:- Srirekha Talasila
Rem ===Input Parameteres: Test Object
'################################################################################################################
 Public Function fn_ObjectExist(obj_TestObject)
 	ObjClass = obj_TestObject.GetTOProperty("micClass")
	Dim int_counter,bln_ObjectExists,Ary_AllItems
	int_counter=0
	int_i=0
	bln_ObjectExists=TRUE	
	Do While NOT (obj_TestObject.exist(0))
		int_counter=int_counter + 1 
		bln_ObjectExists=FALSE
		If int_counter > 500 Then
			fn_ObjectExist=FALSE
			Exit Function 
		End If
	Loop
		
	fn_ObjectExist=TRUE
End Function 

'''#####################################################################################################################
'# Function Name							   -- gfReg_SetOrSelectValue
'#	Description								   -- This function is for set the value or select the value for web objects.
'# Input Parameter	               			   -- obj_TestObject - TestObject ,  str_InputValue - TestData , str_ScreenName - Page details for reporting purpose
'# Designed By                                 -- Srirekha Talasila
''#####################################################################################################################
 Function gfReg_SetorSelect(obj_TestObject,str_InputValue)

	str_Status=fn_ObjectExist(obj_TestObject)
	ObjClass = obj_TestObject.GetTOProperty("micClass")	
	If isnull(str_InputValue) Then				
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"Set or Select","FAIL", "Input Value is not specified .")												
		gfReg_SetorSelect=False
		Exit function
	End If
	If str_Status=TRUE  Then
		
		ObjName = Mid(obj_TestObject.GetROProperty("name"),CINT(INSTRREV(obj_TestObject.GetROProperty("name"),"$"))+1,Len(obj_TestObject.GetROProperty("name")))
		
		If (Ucase(objClass)="WEBEDIT")  OR (Ucase(objClass)="WEBCHECKBOX") Then
			Setting.WebPackage("ReplayType") = 1
			obj_TestObject.set str_InputValue
			Setting.WebPackage("ReplayType") = 2
			Err.Description=str_InputValue & " has been set successfully in the  " & ObjName & "  field"
			str_Status="PASS"
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"Set or Select",str_Status,Err.Description)				
			gfReg_SetorSelect=TRUE
			Exit Function      
		ElseIf (Ucase(objClass)="WEBLIST") OR (Ucase(objClass)="WEBRADIOGROUP")  Then
			obj_TestObject.select str_InputValue
			Err.Description=str_InputValue & " has been set successfully in the  " & ObjName & "  field"
			str_Status="PASS"
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"Set or Select",str_Status,Err.Description)				
			gfReg_SetorSelect=TRUE
			Exit Function   
		End If	
	Else
		Err.Description= ObjClass & chr(32) & " Object not found"
		gfReg_SetorSelect =FALSE
		str_Status="FAIL"		
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"Set or Select",str_Status,Err.Description)		
	End If
	
End Function

'''#####################################################################################################################
'# Function Name							   -- gfReg_Click
'#	Description								   -- This function is for clicks the web objects.
'# Dependencies					  			   -- NA
'# Input Parameter	               			   -- obj_TestObject - TestObject 
'# Designed By                                 -- Srirekha Talasila
'''#####################################################################################################################
 Function gfReg_Click(obj_TestObject)
	
	str_Status=fn_ObjectExist(obj_TestObject)
	ObjClass = obj_TestObject.GetTOProperty("micClass")	
	
	If str_Status=TRUE Then
		   Test_name = Mid(obj_TestObject.GetROProperty("name"),CINT(INSTRREV(obj_TestObject.GetROProperty("name"),"$"))+1,Len(obj_TestObject.GetROProperty("name")))
		   Setting.WebPackage("ReplayType") = 1
		   obj_TestObject.click 		
		   Setting.WebPackage("ReplayType") = 2
		   wait 1
		   Err.Description =  Test_name & chr(32) & "  object found and clicked"											
		   Err.Number = 0										
		   str_Status = "PASS"	
		   gfReg_Click=TRUE		   
		   Call fn_UpdateTestResults(Environment("str_ScreenName"),"Click",str_Status,Err.Description)	
 		   Exit Function	
	ElseIf str_Status=FALSE Then
			Err.Description= ObjClass & chr(32) & " Object not found"
	    	str_Status="FAIL"		
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"Click",str_Status,Err.Description)
	     	gfReg_Click=FALSE
	        Exit Function	
	End If	
	
End Function

'Set Current Row in DataTable

Function SetRow(sheetname)

rowcount= DataTable.GetSheet(sheetname).GetRowCount() 
Set sheetobject=DataTable.GetSheet(sheetname)
For i=1 to rowcount
	sheetobject.SetCurrentRow(i)
	Testcaseid=Datatable.Value("TestcaseID",sheetname)
	If Environment.Value("caseId")=Testcaseid Then
		sheetobject.SetCurrentRow(i)
		Exit For
    End If
Next

End Function

'################################################################################################################
Rem===FunctionName - ExcelReport_Generation
Rem ===Description:- This Function is to Generate the Excel Report
Rem ===Designed By:- Srirekha Talasila
'################################################################################################################

Function ExcelReport_Generation()

	TestDataPath = Environment.Value("ClaimNumberPath") & "\Claim_Numbers\ClaimNumbers.xlsx"
	Set TestData_ExcelObj = CreateObject("Excel.Application")
	TestData_ExcelObj.Workbooks.Open (TestDataPath)
	TestData_ExcelObj.Visible=True
	Set TDSheet = TestData_ExcelObj.Sheets.Item(1)
	RowNum_i=-1
	RowNum_i = TDSheet.usedrange.rows.count  
	TDSheet.Cells(RowNum_i+1,1) =  Environment.Value("SceNum")
	TDSheet.Cells(RowNum_i+1,2) =  Environment.Value("Claim_Number")
	TDSheet.Cells(RowNum_i+1,3) =  Environment.Value("SCaseId")
	TDSheet.Cells(RowNum_i+1,4) =  Date
	TDSheet.Cells(RowNum_i+1,5) =  Environment.value("str_Exe_Status")
	TDSheet.Cells(RowNum_i+1,6) =  Environment.value("str_Exe_Time")
	TestData_ExcelObj.ActiveWorkbook.Save
	TestData_ExcelObj.Application.Quit
	Set TestData_ExcelObj = Nothing
	
End Function

'################################################################################################################
Rem===FunctionName - Update_Regression_Tracker
Rem ===Description:- This Function is to update the ClaimNumber and S-Case in Regression Tracker
Rem ===Designed By:- Srirekha Talasila
'################################################################################################################

Function Update_Regression_Tracker()
	
	Set excelObj = CreateObject("Excel.Application")
	Set excelwrkbook = excelObj.Workbooks.Open(Environment.Value("RegressionTrackerPath"))
	excelObj.Visible = True
	
	Set excelSheet = excelwrkbook.Sheets("Automation Testing")
	Set FindCell = excelSheet.Range("A2:A174").Find(Environment.Value("str_TCID"))
	
	Set excelSheet1 = excelwrkbook.Sheets("XML Validation")
	Set FindCell1 = excelSheet1.Range("A2:A55").Find(Environment.Value("str_TCID"))
	
	Set excelSheet2 = excelwrkbook.Sheets("WC_Stateforms")
	Set FindCell2 = excelSheet2.Range("A3:A8").Find(Environment.Value("str_TCID"))
	
	If NOT FindCell is Nothing Then
		excelSheet.Cells(FindCell.row,2) = Environment.Value("SCaseId")
		excelSheet.Cells(FindCell.row,3) =  Environment.Value("Claim_Number")
		excelObj.Activeworkbook.Save
	
	ElseIf NOT FindCell2 is Nothing  Then
		excelSheet2.Cells(FindCell2.row,2) = Environment.Value("SCaseId")
		excelSheet2.Cells(FindCell2.row,2) = Environment.Value("Claim_Number")
		excelObj.Activeworkbook.Save
	Else
		Set objPopUp = CreateObject("Wscript.Shell")
        objPopUp.Popup Environment.Value("str_TCID") & " Testcase Not Exist in Automation Testing Sheet",1,"Regression Tracker"
        Set objPopUp = Nothing
		Print Environment.Value("str_TCID") & " Testcase Not Exist in Automation Testing Sheet"
	End If
	
	If NOT FindCell1 is Nothing Then
		excelSheet1.Cells(FindCell1.row,2) =  Environment.Value("Claim_Number")
		excelObj.Activeworkbook.Save
	End  IF
	
	excelObj.Application.Quit
	Set FindCell = Nothing
	Set excelSheet = Nothing
	Set excelwrkbook = Nothing
	
End Function

'################################################################################################################
Rem===FunctionName - Clear_Cookies
Rem ===Description:- This Function is to Clear the browser Cookies
Rem ===Designed By:- Srirekha Talasila
'################################################################################################################

Function Clear_Cookies()
	
	SystemUtil.Run "Control.exe","inetcpl.cpl"
	Dialog("text:=Internet Properties").WinButton("text:=&Delete...").Click
	Dialog("text:=Internet Properties").Dialog("text:=Delete Browsing History").WinButton("text:=&Delete").Click
	Wait 3
	Dialog("text:=Internet Properties").WinButton("text:=OK").Click
	
End Function
