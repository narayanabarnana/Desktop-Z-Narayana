'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' About : Procedure to Open a HTML File for Report Log
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function fnOpenHtmlFile()
	 Dim l_objReport	'File Object
	 Dim l_objFS		'File System Object	
	 Dim FolderName ' The name of the results folder name
	g_iPass_Count = 0
	g_iFail_Count = 0
	g_sFileName = sScriptName
	g_iImage_Capture = 1
	g_sFileName=strFolderPath
   	g_sScreenName=strFolderPath
   	
	Set l_objFS = CreateObject("Scripting.FileSystemObject")
	Set l_objReport = l_objFS.OpenTextFile(Environment.Value("HTMLResultsPath"), 2, True)
	l_objReport.Write "<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>" 
	l_objReport.Write "<TR COLS=3><TD BGCOLOR=WHITE WIDTH=6%><IMG SRC='\\uszz1schnas01\interoffice\Training\AppLabs\WFM Automation\Documentation\zurichlogo_donotdelete\zurichlogo.PNG'></TD><TD WIDTH=88% BGCOLOR=WHITE ALIGN=CENTER><FONT FACE=VERDANA COLOR=NAVY SIZE=5><H2><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Environment.Value("str_TCName") & " - Execution status report</B></H2></FONT></TD><TD BGCOLOR=WHITE WIDTH=6%><IMG SRC='\\uszz1schnas01\interoffice\Training\AppLabs\WFM Automation\Documentation\zurichlogo_donotdelete\CSCLogo2.jpg'></TD></TR></TABLE>"
    	l_objReport.Write "<TABLE BORDER=1  CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"
	l_objReport.Write "<TR COLS=2> <TD BGCOLOR=#660099 WIDTH=50% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Machine executed on:"  & "</B></FONT></TD><TD BGCOLOR=#660099 WIDTH=50% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>" & Environment.Value("LocalHostName") & "</B></FONT></TD></TR>"
    	l_objReport.Write "</TABLE></BODY></HTML>"	

		wfmURL=Environment.Value("CP_URL")
		databaseName=Environment.value("Current_Region")
	l_objReport.Write "<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"
	l_objReport.Write "<TR><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B> CP URL:</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=45% COLSPAN=5 ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B>" & wfmURL & "</B></FONT></TD></TR>"
	l_objReport.Write "<TR COLS=4><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B> ENVIRONMENT:</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=35% COLSPAN=5 ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B>" & databaseName & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B>DATE EXECUTED:</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=35% COLSPAN=5 ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B>" & DATE & "</B></FONT></TD></TR>"
	l_objReport.Write "</TABLE></BODY></HTML>"  
	
	l_objReport.Write "<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"	
	l_objReport.Write "<TR COLS=5><TD BGCOLOR=#FFCC99  WIDTH=10%  ALIGN=CENTER><FONT FACE=VERDANA  COLOR=BLACK SIZE=2><B>SL. NUM</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B> SCREEN NAME</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=30% ALIGN=CENTER><FONT FACE=VERDANA COLOR=BLACK SIZE=2 ALIGN=CENTER><B>OPERATION</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=10% ALIGN=CENTER><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>STEP STATUS</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=30% ALIGN=CENTER><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B> STEP DESCRIPTION</B></FONT></TD></TR>"
	
	l_objReport.Close
	Set l_objFS = Nothing
	Set l_objReport = Nothing
	Environment.Value("g_tStart_Time") = Now()
	fnOpenHtmlFile=TRUE
End Function
'*****************************************************************************************************************************************************************************************************************


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' About : Procedure to Open a HTML File for Report Log
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function fnOpenBatchRunHtmlFile(str_HTMLFilePath)
	 Dim l_objReport1	'File Object
	 Dim l_objFS1		'File System Object	
	 Dim FolderName ' The name of the results folder name
	g_iPass_Count = 0
	g_iFail_Count = 0
	g_sFileName = sScriptName
	g_iImage_Capture = 1
	g_sFileName=strFolderPath
   	g_sScreenName=strFolderPath
   	   	  	
	Set l_objFS1 = CreateObject("Scripting.FileSystemObject")
	Set l_objReport1 = l_objFS1.OpenTextFile(str_HTMLFilePath, 2, True)
	l_objReport1.Write "<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>" 
	l_objReport1.Write "<TR COLS=3><TD BGCOLOR=WHITE WIDTH=6%><IMG SRC='\\uszz1schnas01\interoffice\Training\AppLabs\WFM Automation\Documentation\zurichlogo_donotdelete\zurichlogo.PNG'></TD><TD WIDTH=88% BGCOLOR=WHITE ALIGN=CENTER><FONT FACE=VERDANA COLOR=NAVY SIZE=5><H2><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CP Batch Execution status report</B></H2></FONT></TD><TD BGCOLOR=WHITE WIDTH=6%><IMG SRC='\\uszz1schnas01\interoffice\Training\AppLabs\WFM Automation\Documentation\zurichlogo_donotdelete\CSCLogo2.jpg'></TD></TR></TABLE>"
    	l_objReport1.Write "<TABLE BORDER=1  CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"
	l_objReport1.Write "<TR COLS=2> <TD BGCOLOR=#660099 WIDTH=50% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Machine executed on:"  & "</B></FONT></TD><TD BGCOLOR=#660099 WIDTH=50% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>" & Environment.Value("LocalHostName") & "</B></FONT></TD></TR>"
    	l_objReport1.Write "</TABLE></BODY></HTML>"	
    	
    CP_URL=Environment.value("CP_URL")
	
	l_objReport1.Write "<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"
	'l_objReport1.Write "<TR><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B> CP URL:</B></FONT></TD></TR>"
	l_objReport1.Write "<TR COLS=4><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B> ENVIRONMENT:</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=35% COLSPAN=5 ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B>QA</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B>DATE EXECUTED:</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=35% COLSPAN=5 ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B>" & DATE & "</B></FONT></TD></TR>"
	l_objReport1.Write "</TABLE></BODY></HTML>"  
	
	l_objReport1.Write "<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"	
	l_objReport1.Write "<TR COLS=7><TD BGCOLOR=#FFCC99  WIDTH=5%  ALIGN=CENTER><FONT FACE=VERDANA  COLOR=BLACK SIZE=2><B>SCE NUM</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B> CLAIM NUMBER</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=BLACK SIZE=2 ALIGN=CENTER><B>TOTAL STEPS</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>STEPS NOT EXECUTED</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B> PASSED STEPS</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B> FAILED STEPS</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B> EXECUTION TIME</B></FONT></TD></TR>"
	
	l_objReport1.Close
	Set l_objFS1 = Nothing
	Set l_objReport1 = Nothing
	Environment.Value("g_tStart_BatchTime") = Now()
	fnOpenBatchRunHtmlFile=TRUE
End Function
'*****************************************************************************************************************************************************************************************************************

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' About : Procedure to Insert a Section to Report Log
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  Public Function fnInsertSection_BatchRun(str_HTMLFilePath,str_SceHTMLFile)
  	Dim l_objReport1	'File Object
	Dim l_objFS1		'File System Object
	Param1=Environment.Value("SceNum") 
	Param2=Environment.Value("NewClaimNumber")
	Param3=Environment.value("TotalSteps")
	Param4=Environment.value("StepsRemaining")
	Param5=Environment.value("PassedSteps")
	Param6=Environment.value("FailedSteps")
	Param7=Environment.value("str_ExecutionTime")	
	l_sFile=str_SceHTMLFile
	
	Set l_objFS1 = CreateObject("Scripting.FileSystemObject")
	Set l_objReport1 = l_objFS1.OpenTextFile(str_HTMLFilePath, 8, True)		
	'l_objReport.Write "<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"	
	If Cint(Environment.value("FailedSteps"))=0 Then
		l_objReport1.Write "<TR COLS=7><TD BGCOLOR=#EEEEEE  WIDTH=5%  ALIGN=CENTER><FONT FACE=VERDANA  COLOR=GREEN SIZE=2><B>" & Param1 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>" &  Param2 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2 ALIGN=CENTER><B>" &  Param3 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B><A HREF=" & l_sFile & ">" & Param4 & "</A></B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>" &  Param5 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>" &  Param6 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>" &  Param7 & "</B></FONT></TD></TR>"	
		'l_objReport.Write "<TR COLS=5><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=VERDANA SIZE=2>" & sDesc & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=VERDANA SIZE=2>" & sExpected & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=WINGDINGS SIZE=4>2</FONT><FONT FACE=VERDANA SIZE=2><A HREF='" & l_sFile & "'>" & sActual & "</A></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=7%><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=GREEN>P</FONT><FONT FACE=VERDANA SIZE=2 COLOR=GREEN><B>" & sResult & "</B></FONT></TD></TR>"
	ElseIf Cint(Environment.value("FailedSteps"))> 0 Then
		l_objReport1.Write "<TR COLS=7><TD BGCOLOR=#EEEEEE  WIDTH=5%  ALIGN=CENTER><FONT FACE=VERDANA  COLOR=RED SIZE=2><B>" & Param1 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=RED SIZE=2><B>" &  Param2 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=RED SIZE=2 ALIGN=CENTER><B>" &  Param3 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=RED SIZE=2><B><A HREF=" & l_sFile & ">" & Param4 & "</A></B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=RED SIZE=2><B>" &  Param5 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=RED SIZE=2><B>" &  Param6 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=RED SIZE=2><B>" &  Param7 & "</B></FONT></TD></TR>"
		'l_objReport.Write "<TR COLS=5><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=VERDANA SIZE=2>" & sDesc & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=VERDANA SIZE=2>" & sExpected & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=WINGDINGS SIZE=4>2</FONT><FONT FACE=VERDANA SIZE=2><A HREF='" & l_sFile & "'>" & sActual & "</A></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=7%><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=RED>O</FONT><FONT FACE=VERDANA SIZE=2 COLOR=RED><B>" & sResult & "</B></FONT></TD></TR>"
	End If
	
	l_objReport1.Close
	Set l_objFS1 = Nothing
	Set l_objReport1 = Nothing  
	fnInsertSection_BatchRun=TRUE
End Function
'*===========================================================================================================================================================================


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' About : Procedure to Insert a Section to Report Log
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  Public Function fnInsertSection()
  	Dim l_objReport	'File Object
	Dim l_objFS		'File System Object
	'Environment.Value("int_StepNum")=Environment.Value("int_StepNum")+ 1
	Param1=Environment.Value("int_StepNum") & chr(32) & ":" & NOW
	Param2=Environment.Value("str_SheetName")
	Param3=Environment.Value("str_Operation")
	Param4=Environment.Value("str_Status")
	Param5=Environment.Value("ErrDescription")
	l_sFile=Environment.Value("SnapshotsPath")
	Set l_objFS = CreateObject("Scripting.FileSystemObject")
	Set l_objReport = l_objFS.OpenTextFile(Environment.Value("HTMLResultsPath"), 8, True)		
	'l_objReport.Write "<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"	
	If Param4="PASS" Then
'		l_objReport.Write "<TR COLS=5><TD BGCOLOR=#EEEEEE  WIDTH=20%  ALIGN=CENTER><FONT FACE=VERDANA  COLOR=GREEN SIZE=2><B>" & Param1 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>" &  Param2 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2 ALIGN=CENTER><B>" &  Param3 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=10% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B><A HREF=" & l_sFile & ">" & Param4 & "</A></B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=10% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>" &  Param5 & "</B></FONT></TD></TR>"	
		l_objReport.Write "<TR COLS=5><TD BGCOLOR=#EEEEEE  WIDTH=25%  ALIGN=CENTER><FONT FACE=VERDANA  COLOR=GREEN SIZE=2><B>" & Param1 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=25% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>" &  Param2 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2 ALIGN=CENTER><B>" &  Param3 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B><A HREF=" & l_sFile & ">" & Param4 & "</A></B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>" &  Param5 & "</B></FONT></TD></TR>"	
		'l_objReport.Write "<TR COLS=5><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=VERDANA SIZE=2>" & sDesc & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=VERDANA SIZE=2>" & sExpected & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=WINGDINGS SIZE=4>2</FONT><FONT FACE=VERDANA SIZE=2><A HREF='" & l_sFile & "'>" & sActual & "</A></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=7%><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=GREEN>P</FONT><FONT FACE=VERDANA SIZE=2 COLOR=GREEN><B>" & sResult & "</B></FONT></TD></TR>"
	ElseIf Param4="FAIL" Then
		l_objReport.Write "<TR COLS=5><TD BGCOLOR=#EEEEEE  WIDTH=10%  ALIGN=CENTER><FONT FACE=VERDANA  COLOR=RED SIZE=2><B>" & Param1 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=RED SIZE=2><B>" &  Param2 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30% ALIGN=CENTER><FONT FACE=VERDANA COLOR=RED SIZE=2 ALIGN=CENTER><B>" &  Param3 & "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=10% ALIGN=CENTER><FONT FACE=VERDANA COLOR=RED SIZE=2><B><A HREF=" & l_sFile & ">" & Param4 & "</A></B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=10% ALIGN=CENTER><FONT FACE=VERDANA COLOR=RED SIZE=2><B>" &  Param5 & "</B></FONT></TD></TR>"
		'l_objReport.Write "<TR COLS=5><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=VERDANA SIZE=2>" & sDesc & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=VERDANA SIZE=2>" & sExpected & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=25%><FONT FACE=WINGDINGS SIZE=4>2</FONT><FONT FACE=VERDANA SIZE=2><A HREF='" & l_sFile & "'>" & sActual & "</A></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=7%><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=RED>O</FONT><FONT FACE=VERDANA SIZE=2 COLOR=RED><B>" & sResult & "</B></FONT></TD></TR>"
	End If
	
	l_objReport.Close
	Set l_objFS = Nothing
	Set l_objReport = Nothing 	
	fnInsertSection=TRUE
End Function
'*===========================================================================================================================================================================


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' About : Procedure to Close Report Log
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function  fnCloseReport(str_Date)
	Dim l_objReport	'File Object
	Dim l_objFS		'File System Object
	Dim exectime
	Dim strexectime
	Set l_objFS = CreateObject("Scripting.FileSystemObject")
	Set l_objReport = l_objFS.OpenTextFile(Environment.Value("HTMLResultsPath"), 8, True)	
	g_tStart_Time=Environment.Value("g_tStart_Time") 
	g_tEnd_Time = Now()
	Environment.value("TotalSteps")=Environment.Value("int_StepNum")
	Param1=Environment.value("PassedSteps") + Environment.value("FailedSteps")	
	Param2=Environment.value("PassedSteps")
	Param3=Environment.value("FailedSteps")	
	Environment.value("StepsRemaining")=(Cint(Param1)-(cint(Param2)+Cint(Param3)))
	Param4=Environment.value("StepsRemaining")	
	exectime=ROUND((DateDiff("s",g_tStart_Time,g_tEnd_Time))/60,1)
	intExecTime = exectime
	If exectime<1 Then
		exectime=DateDiff("s",g_tStart_Time,g_tEnd_Time)
		strexectime = exectime & "  Sec"
	Else
		exectime=ROUND((DateDiff("s",g_tStart_Time,g_tEnd_Time))/60,1)
		strexectime = exectime & "  Min"
	End If
	Environment.value("str_ExecutionTime")=strexectime
   	l_objReport.Write "<TR COLS=5><TD BGCOLOR=BLACK WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Total Steps : " & Param1 & "</B></FONT></TD><TD BGCOLOR=BLACK WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Steps Not Executed : " & Param4 & "</B></FONT></TD><TD BGCOLOR=BLACK WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Pass Count : " & Param2 & "</B></FONT></TD><TD BGCOLOR=BLACK WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Fail Count : " & Param3 & "</B></FONT></TD><TD BGCOLOR=BLACK WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Total Execution Time : " & strexectime & " </B></FONT></TD></TR>"
	l_objReport.Write "</TABLE></BODY></HTML>"	
	l_objReport.Close
	Set l_objFS = Nothing
	Set l_objReport = Nothing   
	fnCloseHtml=strexectime
End Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' About : Procedure to Close Report Log
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function  fnCloseReport_BatchRun(str_HTMLFilePath)
Dim l_objReport1	'File Object
	Dim l_objFS1		'File System Object
	Dim exectime
	Dim strexectime
	Param1=Environment.value("int_TotalScenarioToExecute")
	Param2=Environment.value("int_TotalPassedScenarios")
	Param3=Environment.value("int_TotalFailedScenarios")	
	Environment.value("int_StepsRemaining")=(Cint(Param1)-(cint(Param2)+Cint(Param3)))
	Param4=Environment.value("int_StepsRemaining")
	
	Set l_objFS1 = CreateObject("Scripting.FileSystemObject")
	Set l_objReport1 = l_objFS1.OpenTextFile(str_HTMLFilePath, 8, True)	
	g_tStart_Time=Environment.Value("g_tStart_BatchTime") 
	g_tEnd_Time = Now()	
		
	exectime=ROUND((DateDiff("s",g_tStart_Time,g_tEnd_Time))/60,1)
	intExecTime = exectime
	If exectime<1 Then
		exectime=DateDiff("s",g_tStart_Time,g_tEnd_Time)
		strexectime = exectime & "  Sec"
	Else
		exectime=ROUND((DateDiff("s",g_tStart_Time,g_tEnd_Time))/60,1)
		strexectime = exectime & "  Min"
	End If
	Environment.value("str_ExecutionTime")=strexectime
   	 l_objReport1.Write "<TR COLS=7><TD BGCOLOR=BLACK WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Total Scenarios : " & Param1 & "</B></FONT></TD><TD BGCOLOR=BLACK WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Scenarios Not Executed : " & Param4 & "</B></FONT></TD> <TD BGCOLOR=BLACK WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B> </B></FONT></TD> <TD BGCOLOR=BLACK WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B></B></FONT></TD><TD BGCOLOR=BLACK WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Passed Scenarios Count : " & Param2 & "</B></FONT></TD> <TD BGCOLOR=BLACK WIDTH=15% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Failed Scenarios Count : " & Param3 & "</B></FONT></TD> <TD BGCOLOR=BLACK WIDTH=25% ALIGN=CENTER><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Total Execution Time : " & strexectime & " </B></FONT></TD></TR>"
	 l_objReport1.Write "</TABLE></BODY></HTML>"	
	l_objReport1.Close
	Set l_objFS1 = Nothing
	Set l_objReport1 = Nothing 
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name   :   Reporting_Initialize
'Description     :   Initializing for the reporting
'----------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Reporting_Initialize()
		Environment.Value("ScreenShotCount")=0
		Call Reporting_CreateFolder()
		Call fnOpenBatchRunHtmlFile(Environment.Value("HTMLSummaryFile"))
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'Sub Name             :   Reporting_CreateFolder
'Description          :   Creates a new Folder for storing the Results based on the present date
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Reporting_CreateFolder()

				Set ObjFSO = CreateObject("Scripting.FileSystemObject")
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				
				Environment.Value("ExecutionResultsPath") =Environment.Value("ExecutionResultsPath") & "\"  & Replace(Date,"/","_")
				If NOT  ObjFSO.FolderExists( Environment.Value("ExecutionResultsPath")) Then
							ObjFSO.CreateFolder(Environment.Value("ExecutionResultsPath"))
				End If
				
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

				str_SystemName=Environment.Value("LocalHostName")
				Environment.Value("ExecutionResultsPath") =Environment.Value("ExecutionResultsPath") & "\"  & str_SystemName
				If NOT  ObjFSO.FolderExists( Environment.Value("ExecutionResultsPath")) Then
							ObjFSO.CreateFolder(Environment.Value("ExecutionResultsPath"))
				End If
				
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++				

				Environment.Value("HTMLSummaryFile") = Environment.Value("ExecutionResultsPath") & "\CP_Execution_Summary" & ".html"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++				
		
		If  ObjFSO.FileExists(Environment.Value("HTMLSummaryFile")) Then
					Environment.Value("NewHTMLBatchResultsPath")=Environment.Value("ExecutionResultsPath") & "\CP_Execution_Summary_" &  Replace(Replace(Replace(NOW,"/","_")," " ,"_"),":","_")  & "Res.htm"	  
					ObjFSO.CopyFile Environment.Value("HTMLSummaryFile"), Environment.Value("NewHTMLBatchResultsPath")
					ObjFSO.DeleteFile Environment.Value("HTMLSummaryFile"),True
		End If	

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				Set ObjFSO = Nothing

		    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name  :   Reporting_TestCase_Initialize
'Description    :   Initializing for the test case reporting
'-----------------------------------------------------------------------

Public Function Reporting_TestCase_Initialize()
		
		Environment.Value("ScreenShotCount")=0
		
		Set ObjFSO = CreateObject("Scripting.FileSystemObject")
		
		Environment.Value("str_ResultsFilePath1")=Environment.Value("ExecutionResultsPath") & "\Scenario_" & Environment.Value("SceNum")
		
		'* Create folder if not exists				
		If NOT  ObjFSO.FolderExists( Environment.Value("str_ResultsFilePath1")) Then
				ObjFSO.CreateFolder(Environment.Value("str_ResultsFilePath1"))
		End If
		
		Environment.Value("HTMLResPath") = Environment.Value("str_ResultsFilePath1") & "\HTML"
		
		If Not ObjFSO.FolderExists(Environment.Value("HTMLResPath")) Then
			ObjFSO.CreateFolder(Environment.Value("HTMLResPath"))
		End If
				
		'Create a screenshot folder with in execution folder
		Environment.Value("ScreenShotPath")=Environment.Value("str_ResultsFilePath1")&"\SCREENSHOTS"
		If Not ObjFSO.FolderExists(Environment.Value("ScreenShotPath")) Then
			 ObjFSO.CreateFolder ( Environment.Value("ScreenShotPath"))
		End If
		
		'Creating a HTML log for a test case which we would be executed
		
		Environment.Value("HTMLResultsPath") = Environment.Value("HTMLResPath") & "\" & Environment.Value("SceNum") & ".html"
		If  ObjFSO.FileExists(Environment.Value("HTMLResultsPath")) Then
					Environment.Value("NewHTMLResultsPath")= Environment.Value("HTMLResPath") & "\" & Environment.Value("SceNum") &"_" &  Replace(Replace(Replace(NOW,"/","_")," " ,"_"),":","_")  & "Res.htm"	  
					ObjFSO.CopyFile Environment.Value("HTMLResultsPath"), Environment.Value("NewHTMLResultsPath")
					ObjFSO.DeleteFile Environment.Value("HTMLResultsPath"),True
		End If	

		Set ObjLogFile = ObjFSO.CreateTextFile(Environment.Value("HTMLResultsPath"), True)
		
		ObjLogFile.Close
		
		Print Environment.Value("HTMLSummaryFile")
		Print Environment.Value("HTMLResultsPath") 
		
		Call fnOpenHtmlFile()
		Set ObjFSO = Nothing
End Function
