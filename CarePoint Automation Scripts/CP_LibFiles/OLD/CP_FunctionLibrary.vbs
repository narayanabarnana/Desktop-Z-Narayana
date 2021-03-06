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
Rem ===Designed By:- CP Automation Team
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
Rem ===Designed By:- CP Automation Team
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
		Wait 1
		bln_ObjectExists=FALSE
		If int_counter>500 Then
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

