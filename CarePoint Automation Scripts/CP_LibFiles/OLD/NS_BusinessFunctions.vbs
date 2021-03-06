	

'General Header
'#####################################################################################################################
'Test Tool/Version		: Quick Test Professional 9.2/9.5/10
'Test Tool Settings		: N.A.
'Browser used			: N.A.
'Application Automated		: 
'Test Case Automated		: N.A.
'Script Name			: Business Components
'Author				: 
'#####################################################################################################################
'Option Explicit 'Forcing Variable declarations

Dim blnverify, Item_count, i, Accident_SiteAddr, pol_flag, Claim_Number,ez_flag,firstpp_flag,peril_flag,bolier_flag,EzRegStatus,ForFirstClaimnt
Dim funarr(100)
ez_flag = False
firstpp_flag = False
peril_flag = False
boiler_flag = False

 Function Login()

	EzRegStatus = "False"
	Systemutil.Run "iexplore.exe", Environment.Value("CP_URL")	
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("UserIdentifier").Set Environment.Value("CP_LoginId")
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("Password").Set Environment.Value("CP_LoginPassword") 
	Browser("ClaimsBrowser").Page("LoginPage").WebButton("Log In").Click
	funarr(1) = True
	
 End function
'######################################################################################################################

'General Header
'#####################################################################################################################
'Test Tool/Version		: Quick Test Professional 9.2/9.5/10
'Test Tool Settings		: N.A.
'Browser used			: N.A.
'Application Automated		: 
'Test Case Automated		: N.A.
'Script Name			: Business Components
'Author				: 
'Date Created			: 
'Last Modified by		: 
'Date Modified			: 
'Comments			: 
'#####################################################################################################################

Function Select_NS()

	If  funarr(1) = True Then
		 ReportResult_Event micPass, "Invoking Business component: Login" , "Invoking Business component: Login - Done"
	Else
		ReportResult_Event micFail, "Invoking Business component: Login" , "Invoking Business component: Login - Failed"
		Excel_Comments = Excel_Comments & "* Invoking Business component: Login - Failed *"
	End if
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1_2").WebElement("My Group").Click
	Wait(2)
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1_2").WebList("select").Select "Non-Standard"
	funarr(2) = True
	
End function

'######################################################################################################################
'General Header
'#####################################################################################################################
'Test Tool/Version		: Quick Test Professional 9.2/9.5/10
'Test Tool Settings		: N.A.
'Browser used			: N.A.
'Application Automated		: 
'Test Case Automated		: N.A.
'Script Name			: Business Components
'Author				: 
'Date Created			: 
'Last Modified by		: 
'Date Modified			: 
'Comments			: 
'#####################################################################################################################

Function Select_Workitem()

	Browser("ACT II").Page("get worklist for selected").WebElement("class:=lv_header_col","html id:=yui-gen8","html tag:=DIV").Click
	Wait(3)
	SelectionCount=1
	Do
		If SelectionCount=1 Then
			Set tabobj=Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection")
			rowcount=trim(tabobj.GetROProperty("rows"))
			For row=2 To rowcount 
				Set tabobj=Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection")
				Status=Trim(tabobj.GetCellData(row,3))
				Channel=Trim(tabobj.GetCellData(row,4))
				IncidentID=Trim(tabobj.GetCellData(row,8))	
				IDType=left(IncidentID,1)
				currentrowcount=row
			'''*****************************************************************************************
				If row=13 Then
				   Set obj = CreateObject("WScript.Shell")
				   obj.SendKeys ("{PGDN}")
				   Set obj=nothing 	
				End If
			'''***************************************************************************************** 
				If Status="New"  and IDType<>"S"  and Channel <> "WEB" and  Channel<>"FTP"  Then
	     			Set objref=createobject("Mercury.DeviceReplay")
				    x=Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").ChildItem(row,3,"WebElement",0).GetRoProperty("abs_x")
				    y=Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").ChildItem(row,3,"WebElement",0).GetRoProperty("abs_y")
				    objref.MouseDblClick x,y,0   
				    Set objref=nothing   
					Exit For			    
				End If		
				If  row=Cint(rowcount) Then
					CustomerSearchCheck=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Customer Search").Exist
					If  CustomerSearchCheck=False Then
						Browser("name:=get worklist for selected workbasket ID").Page("title:=get worklist for selected workbasket ID").Link("name:=Next").Click
						row=1
						Wait 3
					End If
				End If
			Next
			Set tabobj=nothing
			SelectionCount=SelectionCount+1	
		Else
			'Do nothing
		End If
	Check=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Customer Search").Exist
	If Check= "True" Then
		Exit Do
	End If
	Loop Until Check=False		
 
	If  funarr(2) = True Then
		ReportResult_Event micPass, "Invoking Business component: Select_GL" , "Invoking Business component: Select_GL - Done"
	Else
		ReportResult_Event micFail, "Invoking Business component: Select_GL" , "Invoking Business component: Select_GL - Failed"
		Excel_Comments = Excel_Comments & "* Invoking Business component: Select_GL - Failed *" 
	End if
	
	
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Customer Search").Exist Then
		If DataTable("Search_Flow","GL-Data") = "Customer" Then 
			Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Customer Search").Click 
			Wait(7)
			Customer_Search()
		Else
			Employee_Search()
		End if
		ReportResult_Event micPass, "GL Workitem is found in the work queue" , "GL Workitem is found in the work queue - Done"
	Else
		ReportResult_Event micFail, "GL Workitem  in the work queue" , "GL Workitem is not found in the work queue "
		Excel_Comments = Excel_Comments & "* GL Workitem is not found in the work queue *"
	End if		
	

End Function
'#####################################################################################################################


'General Header
'#####################################################################################################################
'Test Tool/Version		: Quick Test Professional 9.2/9.5/10
'Test Tool Settings		: N.A.
'Browser used			: N.A.
'Application Automated		: 
'Test Case Automated		: N.A.
'Script Name			: Business Components
'Author				: 
'Date Created			: 
'Last Modified by		: 
'Date Modified			: 
'Comments			: 
'#####################################################################################################################

Function Customer_Search()
 
	Dim objBrwpage_CustomerSearch
	
	call  SetRow("GL-Data")
	Set objBrwpage_CustomerSearch=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame")
	objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Set DataTable("CS_Accident_Date","GL-Data")	
	objBrwpage_CustomerSearch.WebEdit("CS_CustomerName").Set DataTable("CS_Customer_Name","GL-Data")
	Wait(1)
	objBrwpage_CustomerSearch.WebEdit("CS_SiteCode").Set DataTable("CS_SiteCode","GL-Data")
	Wait(1)
	val=objBrwpage_CustomerSearch.WebEdit("CS_SiteCode").GetROProperty("value")
	If Val<>""  Then
		objBrwpage_CustomerSearch.WebEdit("CS_SiteCode").Set ""
	End If
	objBrwpage_CustomerSearch.WebButton("CS_Search").Click		
	Wait(8)
	Index=1
	While index<>0'''Here the condition will Waits till Web Table load
		index=0
		Set obj_BusinessUnit=Browser("name:=.*Manager.*").Page("title:=.*Manager.*").Frame("name:=actionIFrame").WebTable("column names:=Click to sortBusiness Unit ,;Click to sortCustomer Name ,;Click to sortEntity Name ,;Click to sortSite Name ,;Click to sortSite Code ,;Click to sortAddress 1 ,;Click to sortAddress 2 ,;Click to sortCity ,;Click to sortState ,;Click to sortZip Code ,;Click to sortPhone ,;Click to sortFax ,","index:=23").ChildItem(2,1,"WebElement",0)''@DP
		obj_BusinessUnit.click '''This will target first row in the Customer SEarch result 
		Wait(2)
		objBrwpage_CustomerSearch.WebButton("CS_Select").Click
		Wait(2)
	
		If Browser("name:=CCC Manager Portal 7.1").Dialog("regexpwndtitle:=Message from webpage.*").WinButton("regexpwndtitle:=OK").Exist then
			Browser("name:=CCC Manager Portal 7.1").Dialog("regexpwndtitle:=Message from webpage.*").WinButton("regexpwndtitle:=OK").Click	
		End if
	Wend
	
	If Browser("name:=CCC Manager Portal 7.1").Page("title:=CCC Manager Portal 7.1").Frame("name:=actionIFrame").WebButton("name:=Start Process").Exist Then
		Browser("name:=CCC Manager Portal 7.1").Page("title:=CCC Manager Portal 7.1").Frame("name:=actionIFrame").WebButton("name:=Start Process").click 
		If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist then
			Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
		End If
	End If
	
	If Browser("name:=Care.*").Exist Then
	   	Browser("name:=Care.*").Close 		   
	Else 
		If Browser("name:=http.*").Exist Then
			Browser("name:=http.*").Close 
		End If  
	End If 	

End Function
'#####################################################################################################################

'General Header
'#####################################################################################################################
'Test Tool/Version		: Quick Test Professional 9.2/9.5/10
'Test Tool Settings		: N.A.
'Browser used			: N.A.
'Application Automated		: 
'Test Case Automated		: N.A.
'Script Name			: Business Components
'Author				: 
'Date Created			: 
'Last Modified by		: 
'Date Modified			: 
'Comments			: 
'#####################################################################################################################

Function Employee_Search()

	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Employee Search").Click
	Wait 5
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Emp_CustomerName").Exist Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Emp_CustomerName").Select DataTable("Emp_CustomerName","GL-Data")
		Wait(1)
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Emp_Search").Click
		Wait 4
	End If
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebRadioGroup("Emp_Result").Exist Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebRadioGroup("Emp_Result").Select "1"
	End If
    Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Emp_Select").Click
	Wait(1)
	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("actionIFrame").WebButton("Start Process").Exist Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("actionIFrame").WebButton("Start Process").Click
		Wait 3
	End If
	Wait 3
	If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist then
		Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
	End If 

End Function

'#####################################################################################################################
Function AddCustomer()

		Wait(2)
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Add Customer").Click
		Wait(2)
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Name").Set DataTable("AddCust_Name","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Address").Set DataTable("AddCust_Address","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Zip").Set DataTable("AddCust_Zip","GL-Data")
		Wait(2)
		'Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Phone").Set DataTable("AddCust_Phone","Property")
		'Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Fax").Set DataTable("AddCust_Fax","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Email").Set DataTable("AddCust_Email","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("EmployerTaxID").Set DataTable("AddCust_EmployerTaxID","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("AddCust_Submit").Click

End Function

Function Extract_SCaseId ()
	
	SCase_Id=""
	SCase_Id = Trim(Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("SCaseId").GetROProperty("innertext"))
	
	Environment.Value("SCaseId") = trim(SCase_Id)
	
End Function

Function Incident()
	
	
	Set NS_Incident=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")  
	NS_Incident.WebEdit("IN_AccidentDate").Set DataTable("CS_Accident_Date","GL-Data")
	Wait(1)
	NS_Incident.WebList("AccidentState").Select DataTable("IN_AccidentState","GL-Data")
	Wait(1)
	NS_Incident.WebList("AccidentTime1").Select DataTable("IN_AccidentTime1","GL-Data")
	Wait(1)
	NS_Incident.WebList("AccidentTime2").Select DataTable("IN_AccidentTime2","GL-Data")
	Wait(1)
	NS_Incident.WebList("AccidentTime3").Select DataTable("IN_AccidentTime3","GL-Data")
	Wait(1)
	NS_Incident.WebList("Catagory").Select DataTable("IN_Category","GL-Data")
	Wait(1)
	NS_Incident.WebList("IN_Product").Select DataTable("IN_Product","GL-Data")
	Wait(1)
	If  NS_Incident.WebList("IN_Exposure").Exist Then
		NS_Incident.WebList("IN_Exposure").Select DataTable("IN_Exposure","GL-Data")
	End If
	If DataTable("IN_Product","GL-Data") = "Reinsurance" or DataTable("IN_Product","GL-Data") = "SAFE" or DataTable("IN_Product","GL-Data") = "Occupational Accident"  Then 
	 	NS_Incident.WebEdit("IN_DateReported").Set DataTable("IN_DateReported","GL-Data")
		Wait(1)
	End If
	If  DataTable("IN_Product","GL-Data") = "Surety/Fidelity" or DataTable("IN_Product","GL-Data")="Jockey" Then
		NS_Incident.WebEdit("IN_DateReported").Set DataTable("IN_DateReported","GL-Data")
		Wait(1)
		NS_Incident.WebCheckBox("Is_9_Series").Set DataTable("Is_9_Series","GL-Data")
	End If
	NS_Incident.WebEdit("IN_Claimant_Fname").Set DataTable("IN_Fname","GL-Data")
	Wait(1)
	NS_Incident.WebEdit("IN_Claimant_MI").Set DataTable("IN_MI","GL-Data") 
	Wait(1)
	NS_Incident.WebEdit("IN_Claimant_Lname").Set DataTable("IN_Lname","GL-Data")
	Wait(1)
	NS_Incident.WebEdit("AccidentDescription").Set DataTable("IN_AccDescription","GL-Data")
	Wait(1)
	If DataTable("IN_Product","GL-Data")="Jockey" Then
		Browser("ClaimsBrowser").Page("CCC Bus Admin Portal 7.1").Frame("Frame").WebButton("Save").Click
		Wait(1)
		Browser("ClaimsBrowser").Page("CCC Bus Admin Portal 7.1").Frame("Frame").Image("Jockey - Policy and Address").Click
		Wait(1)
		Window("Windows Internet Explorer").Dialog("Windows Internet Explorer").WinButton("Open").Click
	End If
	
	NS_Incident.WebButton("Next>>").Click
	If  NS_Incident.WebButton("No Duplicates Found").Exist Then 'If Duplicate Claim Exists
		Wait(2)
		 NS_Incident.WebButton("No Duplicates Found").Click
	Else 
		'Do Nothing
	End If
	If Browser("ClaimsBrowser").Dialog("Use_HC_PolicyOnly").Exist then
		Browser("ClaimsBrowser").Dialog("Use_HC_PolicyOnly").WinButton("OK").Click
	End If
	funarr(4) = True
	If   NS_Incident.WebTable("Policy_Table").Exist Then
		ReportResult_Event micPass, "Invoking Business component: Incident" , "Invoking Business component: Incident - Done"
	Else
		ReportResult_Event micFail, "Invoking Business component: Incident" , "Invoking Business component: Incident - Failed"
		Excel_Comments = Excel_Comments & "* Invoking Business component: Incident - Failed *" 
	End if
	
End Function
 
Function PolicySearch()

	If DataTable("IN_Product","GL-Data") = "Occupational Accident"  Then
	   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PS_PolicyNumber").Set DataTable("PS_PolicyNo","GL-Data")
	   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Policy_Retrieve").Click
	   Wait(2)
	   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("Policy_RadioButton").Click
	   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	Else	
		If DataTable("CS_Policynum","GL-Data")<>"" Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PS_PolicyNumber").Set DataTable("CS_Policynum","GL-Data")	
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Policy_Retrieve").Click
		End If 	
			
			Wait(2)
			cell_data = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebTable("Policy_Table").GetCellData(2,1)
			If cell_data = "" Then
				Set polobj = browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebTable("Policy_Table")
				Set polobj2 = polobj.ChildItem(2,1,"WebRadioGroup",0)
				d = polobj2.getroproperty("class")
				If d = "Radio lvInputSelection" Then
					Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("Policy_RadioButton").Click
					Wait(1)
					
					Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
					Wait(2)
				End if
			Else
				''Do Nothing 
			End If			
		End if
	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebElement("PS_NOMatchingData").Exist Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("Policy_RadioButton").Select "Indeterminate"
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Wait(2)
	End If
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Click
	End if
	Wait(2)
	funarr(5) = True
	
End Function

Function Override_TPA()

	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Exist then
		Wait(2)
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Click
	Else
		'Do Nothing
	End If
	
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("CO_Rep_Name").Exist  Then
		ReportResult_Event micPass, "Invoking Business component: PolicySearch" , "Invoking Business component: PolicySearch - Done"
	Else
		ReportResult_Event micFail, "Invoking Business component: PolicySearch" , "Invoking Business component: PolicySearch - Failed"
		Excel_Comments = Excel_Comments & "* Invoking Business component: PolicySearch - Failed *" 
	End if	

End Function

Function Contact_Info()

	Set NS_ConInfo=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 'updated by Rajeshwar on 07/15/2015
	NS_ConInfo.WebEdit("CO_Rep_Name").Set DataTable("CO_Rep_Name","GL-Data")
	NS_ConInfo.WebEdit("CO_Rep_Email").Set DataTable("CO_Rep_Email","GL-Data")
	Wait(2)
	NS_ConInfo.WebEdit("CO_Rep_Phone").Set DataTable("CO_Rep_Phone","GL-Data")
	NS_ConInfo.WebList("CO_Report_Relation").Select DataTable("CO_Rep_Relation","GL-Data")
	'Browser("Inbox").Page("Inbox").Frame("DIACTION").WebEdit("CO_CusCon_Name").Set CRAFT_GetData("CO_CusCon_Name")
	NS_ConInfo.WebEdit("CO_CusCon_Email").Set DataTable("CO_CusCon_Email","GL-Data")
	NS_ConInfo.WebEdit("CO_CusCon_Fax").Set DataTable("CO_CusCon_Fax","GL-Data")		
	NS_ConInfo.WebButton("Next>>").Click
	Wait(4)
	If  NS_ConInfo.WebList("ACC_AccCode").Exist Then
		funarr(6) = True
	Else
		funarr(6) = False
	End If			
	If   NS_ConInfo.WebList("ACC_AccCode").Exist Then
		ReportResult_Event micPass, "Invoking Business component: Contact_Info" , "Invoking Business component: Contact_Info - Done"
	Else
		ReportResult_Event micFail, "Invoking Business component: Contact_Info" , "Invoking Business component: Contact_Info - Failed"
		Excel_Comments = Excel_Comments & "* Invoking Business component: Contact_Info - Failed *" 
	End if

End function
'#####################################################################################################################

Function Accident_Page()

	Set NS_Accident=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")  
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Click
	If DataTable("IN_Product","GL-Data")="Jockey" Then
	   If  NS_Accident.WebList("ACC_AccCode").GetROProperty("default value")="Misc: Unclassified injury" Then 
		   ReportResult_Event micPass, "Invoking Business component: Accident_Page" , "Invoking Business component: Accident_Page - Done"
		Else
			ReportResult_Event micFail, "Invoking Business component: Accident_Page" , "Invoking Business component: Accident_Page - Failed"
			Excel_Comments = Excel_Comments & "* Invoking Business component: Accident_Page - Failed *" 
	   End If 
	Else
		NS_Accident.WebList("ACC_AccCode").Select DataTable("ACC_AccCode","GL-Data")
		Wait(3)
		NS_Accident.WebList("ACC_AgentLoss").Select DataTable("ACC_AgentLoss","GL-Data")
		Wait(3)
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Click
		Wait(3)
		NS_Accident.WebList("ACC_LossLoc").Select DataTable("ACC_LossLoc","GL-Data")
	End If
	
	If DataTable("IN_Product","GL-Data") = "Occupational Accident" Then '''Occupational Accident 
		rem If Browser("name:=CCC Manager Portal 7.1").Page("title:=CCC Manager Portal 7.1").Frame("title:=FNOL.*").WebElement("html tag:=LABEL","innertext:=Please select the state where the driver lives.*").Exist Then
		   Browser("name:=CCC Manager Portal 7.1").Page("title:=CCC Manager Portal 7.1").Frame("title:=FNOL.*").WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAccident.*pBenState","index:=1").Select "DBA"  ''DataTable("ACC_BenefitState","GL-Data")
           ReportResult_Event micPass	, "Validating Benefit State Message: Accident_Page" , "Please select the state where the driver lives Message exist Succesfully for Occupational Accident: Accident_Page - Done "     
        rem Else
        rem   ReportResult_Event micFail, "Validating Benefit State Message: Accident_Page" , "Please select the state where the driver lives Message not exist for Occupational Accident: Accident_Page - Failed"     
		rem End If
	End If
	If  DataTable("ACC_SiteAddress","GL-Data")="No" Then
	 	NS_Accident.WebEdit("ACC_AccAddress1").Set DataTable("ACC_AccAddress1","GL-Data")
	 	NS_Accident.WebEdit("ACC_AccAddress2").Set DataTable("ACC_AccAddress2","GL-Data")
	 	NS_Accident.WebEdit("ACC_AccZip").Set DataTable("ACC_AccZip","GL-Data")
	Else
		NS_Accident.WebList("ACC_SiteAddress").Select DataTable("ACC_SiteAddress","GL-Data")
		Wait(2)
	End If
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("title:=FNOL.*").WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAccident.*pAddr.*pState").GetROProperty("value")="Select..." Then ''''no value exist in the Zip code
		Browser("ClaimsBrowser").Page("Inbox").Frame("title:=FNOL.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAccident.*pAddr.*ppostalCode").Set DataTable("ACC_AccZip","GL-Data")
		ReportResult_Event micPass, "Is the accident address the same as the Partys address : Accident_Page" , "Zip Code Not fetched, So entered Manually " &DataTable("ACC_AccZip","GL-Data")&" entered Manually: Accident_Page - Warning"
		Wait(2)
	End If
	
	If DataTable("IN_Product","GL-Data") = "Reinsurance" or DataTable("IN_Product","GL-Data") = "SAFE" or DataTable("IN_Product","GL-Data") = "Surety/Fidelity" Then
		'do nothing
	Else
		NS_Accident.WebEdit("ACC_Comments").Set DataTable("ACC_Comments","GL-Data")
		' POLICE
		NS_Accident.WebCheckBox("ACC_Police").Set DataTable("ACC_Police","GL-Data")
		NS_Accident.WebCheckBox("ACC_Fire").Set DataTable("ACC_Fire","GL-Data")
		Wait(2)
		NS_Accident.WebCheckBox("ACC_Ambulance").Set DataTable("ACC_Ambulance","GL-Data")
	 	NS_Accident.WebCheckBox("ACC_Other").Set DataTable("ACC_Other","GL-Data")
		If DataTable("ACC_Police","GL-Data") = "ON" Then
			NS_Accident.WebEdit("ACC_Pol_AuthName").Set DataTable("ACC_Pol_AuthName","GL-Data")
			NS_Accident.WebEdit("ACC_Pol_OffName").Set DataTable("ACC_Pol_OffName","GL-Data")
			Wait(2)
			NS_Accident.WebEdit("ACC_Pol_OffBatch").Set DataTable("ACC_Pol_OffBatch","GL-Data")
			NS_Accident.WebEdit("ACC_Pol_Report").Set DataTable("ACC_Pol_Report","GL-Data")
			Wait(2)
			NS_Accident.WebEdit("ACC_Pol_OffPhone").Set DataTable("ACC_Pol_OffPhone","GL-Data")
			Wait(2)
			NS_Accident.WebEdit("ACC_Pol_NCIC").Set DataTable("ACC_Pol_NCIC","GL-Data")
			Wait(2)
		ElseIf ((DataTable("ACC_Fire","GL-Data") = "ON") OR (DataTable("ACC_Ambulance","GL-Data") = "ON") OR (DataTable("ACC_Other","GL-Data") = "ON")) Then
			NS_Accident.WebEdit("ACC_Ambu_AuthName").Set DataTable("ACC_Ambu_AuthName","GL-Data")
			Wait(2)
			NS_Accident.WebEdit("ACC_Ambu_Report").Set DataTable("ACC_Ambu_Report","GL-Data")
			NS_Accident.WebEdit("ACC_Ambu_OSHA").Set DataTable("ACC_Ambu_OSHA","GL-Data")
		End If
	End If
	
	NS_Accident.WebButton("Next>>").Click 
	Wait(2)
	NS_Accident.WebEdit("Party_Fname").WaitProperty "Visible","True",1000
	If   NS_Accident.WebEdit("Party_Fname").Exist Then
		ReportResult_Event micPass, "Invoking Business component: Accident_Page" , "Invoking Business component: Accident_Page - Done"
	Else
		ReportResult_Event micFail, "Invoking Business component: Accident_Page" , "Invoking Business component: Accident_Page - Failed"
		Excel_Comments = Excel_Comments & "* Invoking Business component: Accident_Page - Failed *" 
	End if

End function

'#####################################################################################################################

 
Function Party()

	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Fname").Exist Then
			Dim i
			i = "1"
				counter = Environment.Value("counter")  
				Rem Newly added  code for the "AD & D " Party
					If   DataTable("IN_Product","GL-Data")="AD & D"  Then
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Fname").Set DataTable("Party_Fname","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_MI").Set DataTable("Party_MI","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Lname").Set DataTable("Party_Lname","GL-Data")
							Wait(1)

							If	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Injured").Exist then
								 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Injured").Set DataTable("Party_Injured","GL-Data")
								 Wait(1)
								 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Party_Add_To_List").Click
							 End IF
							 
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Fname").Set DataTable("Party_Fname","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_MI").Set DataTable("Party_MI","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Lname").Set DataTable("Party_Lname","GL-Data")
							Wait(1)

							If	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Injured").Exist then
								 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Injured").Set DataTable("Party_Injured","GL-Data")
								 Wait(1)
								 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Party_Add_To_List").Click
							 End if					
							
								
							If Browser("Inbox").Page("Inbox").Frame("DIACTION").WebElement("3rd Party Injury already").Exist  Then
								Browser("Inbox").Page("Inbox").Frame("DIACTION").WebElement("3rd Party Injury").Click
								Browser("Inbox").Page("Inbox").Frame("DIACTION").WebButton("Delete From List").Click
								Wait(2)
                                Dialog("Message from webpage").WinButton("OK").Click
								Wait(2)
								End IF      
							Rem ================================================

				If DataTable("Different_Claimant_Data","GL-Data") = "Yes" then			
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Fname").Set DataTable("Party_Fname","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_MI").Set DataTable("Party_MI","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Lname").Set DataTable("Party_Lname","GL-Data")
							Wait(1)

							If	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Injured").Exist then
								 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Injured").Set DataTable("Party_Injured","GL-Data")
								 Wait(1)
							End if 
		
							If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Fatality").Exist then
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Fatality").Set DataTable("Party_Fatality","GL-Data")
								Wait(1)
							End if 

							If DataTable("IN_Product","GL-Data")  = "Lawyers Professional Liability" or DataTable("IN_Product","GL-Data")  = "Reinsurance" or DataTable("IN_Product","GL-Data")  = "SAFE" Then
									'do nothing
							Else
									If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_PropertyDamage").Exist then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_PropertyDamage").Set DataTable("Party_PropertyDamage","GL-Data")
										Wait(1)
									End if 
				
									If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Witness").Exist then
										 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Witness").Set DataTable("Party_Witness","GL-Data")
										Wait(1)
									End If
							End If

			End If
							
							For i = 1 to DataTable("No.Of.Claimants","GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Fname").Set DataTable("Party_Fname","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_MI").Set DataTable("Party_MI","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Lname").Set DataTable("Party_Lname","GL-Data")
							Wait(1)

							If	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Injured").Exist then
								 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Injured").Set DataTable("Party_Injured","GL-Data")
								 Wait(1)
							End if 
		
							If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Fatality").Exist then
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Fatality").Set DataTable("Party_Fatality","GL-Data")
								Wait(1)
							End if 

							If DataTable("IN_Product","GL-Data")  = "Lawyers Professional Liability" or DataTable("IN_Product","GL-Data")  = "Reinsurance" or DataTable("IN_Product","GL-Data")  = "SAFE" Then
									'do nothing
							Else
									If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_PropertyDamage").Exist then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_PropertyDamage").Set DataTable("Party_PropertyDamage","GL-Data")
										Wait(1)
									End if 
				
									If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Witness").Exist then
										 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Witness").Set DataTable("Party_Witness","GL-Data")
										Wait(1)
									End If
							End If
							
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Party_Add_To_List").Click
							Wait(1)
							Browser("ClaimsBrowser").Sync
						
											counter =counter + 1
											DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
					Next
				
				Else
				
						  For i = 1 to DataTable("No.Of.Claimants","GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Fname").Set DataTable("Party_Fname","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_MI").Set DataTable("Party_MI","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Party_Lname").Set DataTable("Party_Lname","GL-Data")
							Wait(1)
							
							If	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Injured").Exist then
								 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Injured").Set DataTable("Party_Injured","GL-Data")
								 Wait(1)
							End if 
		
							If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Fatality").Exist then
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Fatality").Set DataTable("Party_Fatality","GL-Data")
								Wait(1)
							End if 
							If DataTable("IN_Product","GL-Data")  = "Lawyers Professional Liability" or DataTable("IN_Product","GL-Data")  = "Reinsurance" or DataTable("IN_Product","GL-Data")  = "SAFE" Then
									'do nothing
							Else
									If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_PropertyDamage").Exist then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_PropertyDamage").Set DataTable("Party_PropertyDamage","GL-Data")
										Wait(1)
									End if 
				
									If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Witness").Exist then
										 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Party_Witness").Set DataTable("Party_Witness","GL-Data")
										Wait(1)
									End If
							End If
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Party_Add_To_List").Click
							Wait(1)
							Browser("ClaimsBrowser").Sync
						Next
		
				End if			
					
	End If


	
			
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
			Wait(3)
		





		If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_Add1").Exist Then
					ReportResult_Event micPass, "Invoking Business component: Party_Page" , "Invoking Business component: Party_Page - Done"
		Else
					ReportResult_Event micFail, "Invoking Business component: Party_Page" , "Invoking Business component: Party_Page - Failed"
					Excel_Comments = Excel_Comments & "* Invoking Business component: Party_Page - Failed *" 
		End If

End function


Function Employment()  ''This function newly added after  Occupational Accident 

	Set Obj_Employment=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pOccupation").Set DataTable("Employee_RegularOccupation","GL-Data")
	wait(2)
	rem Obj_Employment.WebEdit("Emp_Inj_Occupation").Set DataTable("Emp_Injury_Occupation","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pdeptNumber").Set DataTable("Employee_Dept","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pStartDate").Set DataTable("Employee_HireDate","GL-Data")
	Obj_Employment.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pemploymentStatus").Select DataTable("Employee_Status","GL-Data")
	wait(2)
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pSupervisorName").Set DataTable("Employee_SupervisorName","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pTelNbr.*gPhone.*pPhone").Set DataTable("Employee_SupervisorPhone","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pemployerNotifiedDate").Set DataTable("Employee_NotifiedDate","GL-Data")
	wait(2)
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pwageAmt").Set DataTable("Employee_WageAmount","GL-Data")
	Obj_Employment.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*psalaryPaymentFrequency").Select DataTable("Employee_Hourly","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pworkHoursPerDay").Set DataTable("Employee_Hours","GL-Data")
	Obj_Employment.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pworkHoursPerWeek").Select DataTable("Employee_Days","GL-Data")
	Obj_Employment.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pWorkShift").Select DataTable("Employee_WorkShift","GL-Data")
	Obj_Employment.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*plostTimeIndicator").Select DataTable("Employee_LostTime","GL-Data")
	wait(2)
	rem Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*plastDayWorked").Set DataTable("Employee_LDW","GL-Data")
	rem Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pdisabilityDate").Set DataTable("Employee_DisabilityDate","GL-Data")
	rem Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*ppaidThrDate").Set DataTable("Employee_PaidDate","GL-Data")	
	
	 
	
End Function

													
Function PartyInfo1()

		counter = Environment.Value("counter")
		DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
        
		If DataTable("Different_Claimant_Data","GL-Data") = "Yes" then
			For i = 1 to DataTable("No.Of.Claimants","GL-Data")
						Wait(1)
								If i=1 Then
									If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Exist Then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
									End If
								Else
								    Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo2_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No  
                                 End If
								If DataTable("IN_Product","GL-Data")  = "Reinsurance" or DataTable("IN_Product","GL-Data")  = "HBP-HVP" Then
									' DoNothing
								else
									 If i=1 then
										 	call Attorney()
										  Else	'''here i=2/3 i.e in Attorney page objects is unable to identify dude to the over lap of objects in i=1  
												Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("Party2_AttorneyList").Select DataTable("Attorney_List","GL-Data")
												Wait(3)
												If DataTable("Attorney_List","GL-Data") = "Yes" Then
													Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Party2_Att_FirmName").Set DataTable("Attorney_FirmName","GL-Data")
													Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Party2_Att_FirstName").Set DataTable("Attorney_FirstName","GL-Data")
													Wait(2)
													Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Party2_Att_LastName").Set DataTable("Attorney_LastName","GL-Data")
													Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Party2_Att_Address1").Set DataTable("Attorney_Address1","GL-Data")
													Wait(2)
													Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Party2_Att_ZIP").Set DataTable("Attorney_ZIP","GL-Data")
													Wait(2)
												 End If		
												If DataTable("IN_Product","GL-Data") = "Occupational Accident" And Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pEmployeeId").Exist Then
													Call Employment()
												End If 

												If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Exist Then
													funarr(10) = True
													ReportResult_Event micPass, "Invoking Business component: Attorney" , "Invoking Business component: Attorney - Done"
												Else
													funarr(10) = False
													ReportResult_Event micFail, "Invoking Business component: Attorney" , "Invoking Business component: Attorney - Failed"
													Excel_Comments = Excel_Comments & "* Invoking Business component: Attorney - Failed *" 
												End If
									 End If 
								End If
								If DataTable("IN_Product","GL-Data") = "Occupational Accident" And Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pEmployeeId").Exist Then
									Call Employment()
								End If
								If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_Zip").Set DataTable("PartyInfo_Zip","GL-Data")										
								End If
					
								If DataTable("Party_Witness","GL-Data") = "ON" Then
									'do nothing
								Else
                                    Browser("ClaimsBrowser").Sync
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
									Wait(3)
								End if 
						
								
'''''''''''''''''''''''''''''''''''''''''''''''''''Injury  Info1''''''''''''''''''''''''''''''''''''''''''''''''''
							If i=1 Then
								If DataTable("Party_Injured","GL-Data") = "ON" Then
									ForFirstClaimntInj=DataTable("Party_Injured","GL-Data")
									Browser("ClaimsBrowser").Sync
									Wait(1)
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description1").Set DataTable("Inj_Description","GL-Data")
									Wait(1)
									If DataTable("Party_Fatality","GL-Data") = "ON" Then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_DateOfDeath").Set DataTable("Inj_DateOfDeath","GL-Data")
									End If
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_CauseInjury1").Set DataTable("Inj_CauseInjury","GL-Data")
									Wait(1)
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_Nature1").Select DataTable("Inj_Nature","GL-Data")
									Wait(1)
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_BodyPart1").Select  DataTable("Inj_BodyPart","GL-Data")
									Wait(1)
									If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment1").Exist then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment1").Select DataTable("Inj_InitialTreatment","GL-Data")
										Wait(1)
									End if 
									If DataTable("Inj_InitialTreatment","GL-Data") = "NO MEDICAL TREATMENT"  or DataTable("Inj_InitialTreatment","GL-Data") = "MINOR ON-SITE REMEDIES BY EMPLOYER MEDICAL STAFF" Then
							            'do nothing
										else
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_LName").Set DataTable("Inj_Phy_LName", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_FName").Set DataTable("Inj_Phy_FName", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_MI").Set DataTable("Inj_Phy_MI", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Add1").Set DataTable("Inj_Phy_Add1", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Add2").Set DataTable("Inj_Phy_Add2", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Zip").Set DataTable("Inj_Phy_Zip", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Phone").Set DataTable("Inj_Phy_Phone", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Fax").Set DataTable("Inj_Phy_Fax", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Email").Set DataTable("Inj_Phy_Email", "GL-Data")
										Wait(1)
						
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Name").Set DataTable("Inj_Hosp_Name", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Add1").Set DataTable("Inj_Hosp_Add1", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Add2").Set DataTable("Inj_Hosp_Add2", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Zip").Set DataTable("Inj_Hosp_Zip", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Phone").Set DataTable("Inj_Hosp_Phone", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Fax").Set DataTable("Inj_Hosp_Fax", "GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Email").Set DataTable("Inj_Hosp_Email", "GL-Data")
										
									end if	
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
								End If
						Else
							Wait(3)
							DataTable("Party_Injured","GL-Data")=ForFirstClaimntInj
							If DataTable("Party_Injured","GL-Data") = "ON" And Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProInj2_Inj_Description1").Exist Then							
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProInj2_Inj_Description1").Set DataTable("Inj_Description","GL-Data")
									Wait(1)
									If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProInj2_Inj_DateOfDeath").Exist Then''''''''''''''If DataTable("Party_Fatality","GL-Data") = "ON" Then
										Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProInj2_Inj_DateOfDeath").Set DataTable("Inj_DateOfDeath","GL-Data")
									End If
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProInj2_Inj_CauseInjury1").Set DataTable("Inj_CauseInjury","GL-Data")
									Wait(1)
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("ProInj2_Inj_Nature").Select DataTable("Inj_Nature","GL-Data")
									Wait(1)
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("ProInj2_Inj_BodyPart").Select  DataTable("Inj_BodyPart","GL-Data")
									Wait(1)								
									If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("ProInj2_Inj_InitialTreatment1").Exist then
										Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("ProInj2_Inj_InitialTreatment1").Select DataTable("Inj_InitialTreatment","GL-Data")
										Wait(1)
									End if 
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
									Wait(2)
							End If 
						End If 
								''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Property Damage1'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						If i=1 Then
							  If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
							  		ForFirstClaimntDmg=DataTable("Party_PropertyDamage","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PropertyDam1_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
									Wait(2)
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
									Wait(1)
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
									Wait(1)
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
									Wait(1)
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropertyDam1_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
									Wait(1)
									If DataTable("IN_Product","GL-Data") = "Surety/Fidelity" Then
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PD_ClaimType").Select DataTable("PropertyDam_ClaimType","GL-Data")
										Wait(1)
									End if 
											If DataTable("PropertyDam_ClaimType","GL-Data") = "Surety" or DataTable("PropertyDam_ClaimType","GL-Data") = "Fidelity" Then
												Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Name").Set "test"
												Wait(1)
												Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add1").Set "123main st"
												Wait(1)
												Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add2").Set "456main st"
												Wait(1)
												Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Zip").Set "12345"
												Wait(1)
		                                        Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Phone").Set "111-222-3333"
												Wait(1)		
											End if 
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
										Wait(1)
									End If
							Else '''i=2/3   '''In the OR the property Dameage objects in pro1 and prop 2 are overlapped 							
							      DataTable("Party_PropertyDamage","GL-Data")=ForFirstClaimntDmg
							     If DataTable("Party_PropertyDamage","GL-Data") = "ON" and Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProDam2_PropertyDam1_DamDescription").Exist  Then
										 If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebRadioGroup("ProDam2_PropertyDam1_Location").Exist then 
											Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebRadioGroup("ProDam2_PropertyDam1_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
											Wait(2)
										 End If 
										Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProDam2_PropertyDam1_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProDam2_PropertyDam1_DamDescription").Set DataTable("PropertyDam_DamDescription","GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProDam2_PropertyDam1_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("ProDam2_PropertyDam1_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
										Wait(1)
										If DataTable("IN_Product","GL-Data") = "Surety/Fidelity" Then
											Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PD_ClaimType").Select DataTable("PropertyDam_ClaimType","GL-Data")
											Wait(1)
										End if 
											If DataTable("PropertyDam_ClaimType","GL-Data") = "Surety" or DataTable("PropertyDam_ClaimType","GL-Data") = "Fidelity" Then
												Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Name").Set "test"
												Wait(1)
												Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add1").Set "123main st"
												Wait(1)
												Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add2").Set "456main st"
												Wait(1)
												Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Zip").Set "12345"
												Wait(1)
		                                        Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Phone").Set "111-222-3333"
												Wait(1)		
											End if 
										Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebCheckBox("ProDam2_PropertyDam1_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebCheckBox("ProDam2_PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
										Wait(1)
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
										Wait(1)
									End If  
										
							End if 							
									
								If DataTable("Party_Witness","GL-Data") = "ON" Then
									Call Witness()
								End If

								counter = counter + 1
								DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
					Next
			
		Else
			'	For i = 1 to DataTable("No.Of.Claimants","GL-Data")
					counter = Environment.Value("counter")
					DataTable.GetSheet("GL-Data").SetCurrentRow(counter)				
					Wait(1)
					If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Exist Then
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
						Wait(1)
					End If
					If DataTable("IN_Product","GL-Data")  <> "Reinsurance" or DataTable("IN_Product","GL-Data")  = "HBP-HVP" Then
						call Attorney()
					End If						
					Wait(1)
					Browser("ClaimsBrowser").Sync
					If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
						'Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_PriPhone").Set DataTable("PartyInfo_PriPhone","Property")
					End If		
					Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
					Wait(2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Injury  Info1'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					If DataTable("Party_Injured","GL-Data") = "ON" Then
						Browser("ClaimsBrowser").Sync
						Wait(1)
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description1").Set DataTable("Inj_Description","GL-Data")
						Wait(1)
						If DataTable("Party_Fatality","GL-Data") = "ON" Then
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_DateOfDeath").Set DataTable("Inj_DateOfDeath","GL-Data")
						End If				
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_CauseInjury1").Set DataTable("Inj_CauseInjury","GL-Data")
						Wait(1)
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("Inj_Nature1").Select DataTable("Inj_Nature","GL-Data")
						Wait(1)
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("Inj_BodyPart1").Select DataTable("Inj_BodyPart","GL-Data")
						If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("Inj_InitialTreatment1").Exist then
							Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("Inj_InitialTreatment1").Select DataTable("Inj_InitialTreatment","GL-Data")
							Wait(1)
						End if 
						If DataTable("Inj_InitialTreatment","GL-Data") = "NO MEDICAL TREATMENT"  or DataTable("Inj_InitialTreatment","GL-Data") = "MINOR ON-SITE REMEDIES BY EMPLOYER MEDICAL STAFF" Then
							'do nothing
						else						
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_LName").Set DataTable("Inj_Phy_LName", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_FName").Set DataTable("Inj_Phy_FName", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_MI").Set DataTable("Inj_Phy_MI", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Add1").Set DataTable("Inj_Phy_Add1", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Add2").Set DataTable("Inj_Phy_Add2", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Zip").Set DataTable("Inj_Phy_Zip", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Phone").Set DataTable("Inj_Phy_Phone", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Fax").Set DataTable("Inj_Phy_Fax", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Email").Set DataTable("Inj_Phy_Email", "GL-Data")
							Wait(1)			
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Name").Set DataTable("Inj_Hosp_Name", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Add1").Set DataTable("Inj_Hosp_Add1", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Add2").Set DataTable("Inj_Hosp_Add2", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Zip").Set DataTable("Inj_Hosp_Zip", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Phone").Set DataTable("Inj_Hosp_Phone", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Fax").Set DataTable("Inj_Hosp_Fax", "GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Email").Set DataTable("Inj_Hosp_Email", "GL-Data")							
						End if	
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
					End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Injury  Info1'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''				
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Property Damage1''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						Browser("ClaimsBrowser").Sync
						If DataTable("Party_PropertyDamage","GL-Data") = "ON" and Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProDam2_PropertyDam1_DamDescription").Exist  Then
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PropertyDam1_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
							Wait(2)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
							Wait(1)
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropertyDam1_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
							Wait(1)
							If DataTable("IN_Product","GL-Data") = "Surety/Fidelity" Then
								Wait(1)
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PD_ClaimType").Select DataTable("PropertyDam_ClaimType","GL-Data")
								Wait(1)
							End if 
							If DataTable("PropertyDam_ClaimType","GL-Data") = "Surety" or DataTable("PropertyDam_ClaimType","GL-Data") = "Fidelity" Then
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Name").Set "test"
								Wait(1)
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add1").Set "123main st"
								Wait(1)
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add2").Set "456main st"
								Wait(1)
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Zip").Set "12345"
								Wait(1)
                                Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Phone").Set "111-222-3333"
								Wait(1)												
							End if 
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
								Wait(1)
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
								Wait(1)
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
								Wait(1)
						End If
				'	Next
    		End If
    		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Property Damage1''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						counter = Environment.Value("counter")
						DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
End Function

Function PartyInfo2()
		'Party Info 2
'''			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo2_ZurichEmp").Select DataTable("PartyInfo_ZurichEmp","GL-Data")  'value  true/false		
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo2_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
		Wait(3)
		If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo2_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo2_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
			Wait(2)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo2_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
			Wait(2)
			'Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo2_PriPhone").Set DataTable("PartyInfo_PriPhone","Property")
		End if
		
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
        Wait(2)
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Injury 2''''''''''''''''''''''''''''''''''''''''''''
		If DataTable("Party_Injured","GL-Data") = "ON" Then
			Wait(1)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description2").Set DataTable("Inj_Description","GL-Data")
			Wait(1)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_CauseInjury2").Set DataTable("Inj_CauseInjury","GL-Data")
			Wait(1)
			If DataTable("Party_Fatality","GL-Data") = "ON" Then
				Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("ProInj2_Inj_DateOfDeath").Set DataTable("Inj_DateOfDeath","GL-Data")
			End If
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_Nature2").Select DataTable("Inj_Nature","GL-Data")
			Wait(2)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_BodyPart2").Select  DataTable("Inj_BodyPart","GL-Data")
			Wait(2)
			If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment2").Exist then
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment2").Select DataTable("Inj_InitialTreatment","GL-Data")
				Wait(1)
			End if 
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
			Wait(2)
		End  if
		 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Property Damage 2''''''''''''''''''''''''''''''''''''''''''''
		If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
			Wait(1)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PropertyDam2_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
			Wait(2)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam2_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
			Wait(1)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam2_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
			Wait(1)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam2_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
			Wait(1)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropertyDam2_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
			Wait(2)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam2_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
			Wait(1)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam2_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
			Wait(1)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
			Wait(2)
		End if

End Function

Function PartyInfo3()

		'Party Info 3
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo3_ZurichEmp").Select DataTable("PartyInfo_ZurichEmp","GL-Data")  'value  true/false
		Wait(1)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo3_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
		Wait(1)
			Browser("ClaimsBrowser").Sync
		If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo3_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo3_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo3_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
			'Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo3_PriPhone").Set DataTable("PartyInfo_PriPhone","Property")
		End if
		'Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo3_Fax").Set DataTable("PartyInfo_Fax","Property")
		'Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo3_Email").Set DataTable("PartyInfo_Email","Property")
		'Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo3_DOB").Set DataTable("PartyInfo_DOB","Property")
		'Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PartyInfo3_Distribution").Select DataTable("PartyInfo_Distribution","Property")
		Browser("ClaimsBrowser").Sync
		'Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PartyInfo3_Gender").Select DataTable("PartyInfo_Gender","Property")
'		Browser("ClaimsBrowser").Sync
'		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PartyInfo3_Marital").Select DataTable("PartyInfo_Marital","Property")
		Browser("ClaimsBrowser").Sync
		Wait(1)
		
		
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click


		'Injury 3
		If DataTable("Party_Injured","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
			Wait(1)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description3").Set DataTable("Inj_Description","GL-Data")
			Wait(1)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_CauseInjury3").Set DataTable("Inj_CauseInjury","GL-Data")
			Wait(1)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_Nature3").Select DataTable("Inj_Nature","GL-Data")
			Wait(2)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_BodyPart3").Select  DataTable("Inj_BodyPart","GL-Data")
			Wait(2)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment3").Select DataTable("Inj_InitialTreatment","GL-Data")
			Wait(2)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		End if


		'Property Damage 3
		If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
				Browser("ClaimsBrowser").Sync
			Wait(1)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PropertyDam3_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
			Wait(2)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam3_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
			Wait(1)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam3_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
			Wait(1)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam3_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
			Wait(1)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropertyDam3_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
			Wait(2)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam3_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
			Wait(1)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam3_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
			Wait(1)
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		End if 
End function 

'#####################################################################################################################
'General Header
'#####################################################################################################################
'Test Tool/Version		: Quick Test Professional 9.2/9.5/10
'Test Tool Settings		: N.A.
'Browser used			: N.A.
'Application Automated		: 
'Test Case Automated		: N.A.
'Script Name			: Business Components
'Author				: 
'Date Created			: 
'Last Modified by		: 
'Date Modified			: 
'Comments			: 
'#####################################################################################################################


''''#############################

Function Witness()
	
	If DataTable("Party_Witness","GL-Data") = "ON" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Wit_FirstName").Set DataTable("Witness_FirstName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Wit_LastName").Set DataTable("Witness_LastName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Wit_Address1").Set DataTable("Witness_Address1","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Wit_Zip").Set DataTable("Witness_Zip","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=HomePhone","html tag:=INPUT").Set DataTable("Witness_PrimaryPhone","GL-Data")
		Wait(2)
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=Fax","html tag:=INPUT").Set DataTable("Witness_Fax","GL-Data")
		
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Exist Then
			ReportResult_Event micPass, "Invoking Business component: Witness" , "Invoking Business component: Witness - Done"
		Else
			ReportResult_Event micFail, "Invoking Business component: Witness" , "Invoking Business component: Witness - Failed"
			Excel_Comments = Excel_Comments & "* Invoking Business component: Witness - Failed *" 
		End If
	End If

End Function
'#####################################################################################################################

'General Header
'#####################################################################################################################
'Test Tool/Version		: Quick Test Professional 9.2/9.5/10
'Test Tool Settings		: N.A.
'Browser used			: N.A.
'Application Automated		: 
'Test Case Automated		: N.A.
'Script Name			: Business Components
'Author				: 
'Date Created			: 
'Last Modified by		: 
'Date Modified			: 
'Comments			: 
'#####################################################################################################################

Function Attorney()
	
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("AttorneyList").Select DataTable("Attorney_List","GL-Data")
	Wait(2)
	If DataTable("Attorney_List","GL-Data") = "Yes" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_FirmName").Set DataTable("Attorney_FirmName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_FirstName").Set DataTable("Attorney_FirstName","GL-Data")
		Wait(2)
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_LastName").Set DataTable("Attorney_LastName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_Address1").Set DataTable("Attorney_Address1","GL-Data")
		Wait(2)
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_ZIP").Set DataTable("Attorney_ZIP","GL-Data")
	End If
	
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Exist Then
		funarr(10) = True
		ReportResult_Event micPass, "Invoking Business component: Attorney" , "Invoking Business component: Attorney - Done"
	Else
		funarr(10) = False
		ReportResult_Event micFail, "Invoking Business component: Attorney" , "Invoking Business component: Attorney - Failed"
		Excel_Comments = Excel_Comments & "* Invoking Business component: Attorney - Failed *" 
	End If
	 
End Function
''#####################################################################################################################

'General Header
'#####################################################################################################################
'Test Tool/Version		: Quick Test Professional 9.2/9.5/10
'Test Tool Settings		: N.A.
'Browser used			: N.A.
'Application Automated		: 
'Test Case Automated		: N.A.
'Script Name			: Business Components
'Author				: 
'Date Created			: 
'Last Modified by		: 
'Date Modified			: 
'Comments			: 
'#####################################################################################################################

Function Additional_Information()

	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Addinfo_NoticOfOccurance").GetROProperty("width")>0 Then 
		If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("title:=FNOL.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pnoticeOfClaim").GetROProperty("value")= "" Then 
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("title:=FNOL.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pnoticeOfOccurence").Set DataTable("AddInfo_NoticeOfOccurance","GL-Data")
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("title:=FNOL.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pnoticeOfClaim").Set DataTable("AddInfo_NoticeOfClaim","GL-Data")
		End If 
	End If 
	If  Datatable("ZurichAttorney","GL-Data")="Yes"  Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ZurichAttorney").Select  DataTable("ZurichAttorney","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ZurichAttorney_FirmName").Set   DataTable("ZurichAttorney_FirmName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ZurichAttorney_FName").Set  DataTable("ZurichAttorney_FName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ZurichAttorney_LName").Set  DataTable("ZurichAttorney_LName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ZurichAttorney_Address1").Set  DataTable("ZurichAttorney_Address1","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ZurichAttorney_Zip").Set  DataTable("ZurichAttorney_Zip","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ZurichAttorney_Notes").Set  DataTable("ZurichAttorney_Notes","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ZurichAttorney_Category").Select  DataTable("ZurichAttorney_Category","GL-Data")
	End If
	If  Datatable("ZurichAttorney","GL-Data")="No"  Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ZurichAttorney_Notes").Set  DataTable("ZurichAttorney_Notes","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ZurichAttorney_Category").Select  DataTable("ZurichAttorney_Category","GL-Data")
		Wait(2)
	End If
	
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	funarr(10) = true
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Get_Claim_Number").Exist Then
		ReportResult_Event micPass, "Invoking Business component: Additional_Information" , "Invoking Business component: Additional_Information - Done"
	Else
		ReportResult_Event micFail, "Invoking Business component: Additional_Information" , "Invoking Business component: Additional_Information - Failed"
		Excel_Comments = Excel_Comments & "* Invoking Business component: Additional_Information - Failed *" 
	End if	

End Function

Function Assignment()

	If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Run Assignment").Exist Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Accident_Code").Select "#01"' DataTable("Assignment_Acc_Code","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Run Assignment").Click
		Wait(3)
	End If	
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Click
	
	If 	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Get_Claim_Number").Exist Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Get_Claim_Number").Click
		Wait(3)
	End If
	
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Click
	End If
	
	funarr(11) = True
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Dist_Complete").Exist Then
		ReportResult_Event micPass, "Invoking Business component: Assignment" , "Invoking Business component: Assignment - Done"
	Else
		ReportResult_Event micFail, "Invoking Business component: Assignment" , "Invoking Business component: Assignment - Failed"
		Excel_Comments = Excel_Comments & "* Invoking Business component: Assignment - Failed *" 
	End if	
		
End Function

Function Review_Distribution()
	
	If  DataTable("CS_Policynum","GL-Data") = "00656613" Then
		Claim_Number =" Farmer's Policy"
		Environment.Value("ClaimNumber") = Claim_Number
	Else	
		Claim_Number= Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("Claim_Number").GetROProperty("innertext")
		Claim_Number=Trim(Claim_Number)
		Claim_Number=right(Claim_Number,10)
	End If
	Environment.Value("ClaimNumber") = Claim_Number
	
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Dist_Complete").Exist Then
	   ReportResult_Event micPass, "Invoking Business component: Review_Distribution" , "Invoking Business component: Review_Distribution - Done"
	   Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Dist_Complete").Click	
	   Wait(2)	   
	Else
		ReportResult_Event micFail, "Invoking Business component: Review_Distribution" , "Invoking Business component: Review_Distribution - Failed"
		Excel_Comments = Excel_Comments & "* Invoking Business component: Review_Distribution - Failed *" 
	End If		
	funarr(12) = True
	Environment.Value("CarePA_Status") = Environment.Value("TC_Status")	
	

	
 End Function
 
 
Function CarePoint_LoggOff()

	Browser("name:=CCC Manager Portal 7.1").Page("title:=CCC Manager Portal 7.1").Link("name:=Manager.*").Click
	Wait(1)
	Browser("name:=CCC Manager Portal 7.1").Page("title:=CCC Manager Portal 7.1").WebElement("innerhtml:=Log off","class:=middleBack","html id:=ItemMiddle").Click
	 
 End Function

Function Binocular_search()
	
	Dim incidentsearch,EXP_IncidentID
	Browser("ClaimsBrowser").Page("Inbox").WebElement("BinocularSearch").Click
	Wait(3)
	Browser("QA: Zurich Intranet Login").Page("SearchIncident").WebEdit("BinocularSearch_ClaimNbr").Set Environment.Value("ClaimNumber")
	Browser("QA: Zurich Intranet Login").Page("SearchIncident").WebButton("Search").Click	
	Wait(3)
	Set WshShell = CreateObject("WScript.Shell") 
	WshShell.SendKeys "%{ }" 
	Wait(3) 
	WshShell.SendKeys " x" 
	Set WshShell=Nothing 	
	Wait(4)
	Browser("QA: Zurich Intranet Login").Page("SearchIncident").WebElement("Incident ID").Click
	Wait(3)
	rem Browser("QA: Zurich Intranet Login").Page("SearchIncident").WebTable("ClaimNbrSearch_IncidentID").ChildItem(2,1,"WebElement",0).click 	
	Browser("QA: Zurich Intranet Login").Close	
	Wait(2)
	If Trim(Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("Resolved-Completed").GetROProperty("innertext"))="Resolved-Completed" Then 
		'''Do Nothing 
	Else
		EXPCase_Number=Trim(Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Link("EXP-Case").GetROProperty("text"))
       
		rem Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame_2").WebTable("EXP-Case Number").GetCellData(4,1)
	End If 
	Browser("ClaimsBrowser").Page("Inbox").WebElement("Inbox").Click
	Wait(2)
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Queue").Select "Exception Handling"
    Wait(2)
	Browser("ACT II").Page("get worklist for selected").WebElement("SortDate").Click
	Wait(2)
	If  Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").Exist Then
		ReportResult_Event micPass, "Invoking Business component: Exception Handling" , "Invoking Business component: Exception Handling - Done"
	Else
		ReportResult_Event micFail, "Invoking Business component: Exception Handling" , "Invoking Business component: Exception Handling - Failed"		
	End if
	
	EXP_IncidentID=Trim(Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").GetCellData(2,8))
	If EXP_IncidentID=EXPCase_Number Then		
		ReportResult_Event micPass, "Validate Exp Case in WorkItem Selection : Binocular_search" , "Exp Case in WorkItem Selection match with Incident History"
	Else	
		ReportResult_Event micFail, "Validate Exp Case in WorkItem Selection : Binocular_search" , "Exp Case in WorkItem Selection Does Not match with Incident History"		
	End If
	Set objref=createobject("Mercury.DeviceReplay")
    x=Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").ChildItem(2,3,"WebElement",0).GetRoProperty("abs_x")
    y=Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").ChildItem(2,3,"WebElement",0).GetRoProperty("abs_y")
    objref.MouseDblClick x,y,0   
    Set objref=nothing   
	
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Customer Search").Exist Then		 
		Customer_Search()
	End If 
	
End Function
'#####################################################################################################################

'General Header
'#####################################################################################################################
'Test Tool/Version		: Quick Test Professional 9.2/9.5/10
'Test Tool Settings		: N.A.
'Browser used			: N.A.
'Application Automated		: 
'Test Case Automated		: N.A.
'Script Name			: Business Components
'Author				: 
'Date Created			: 
'Last Modified by		: 
'Date Modified			: 
'Comments			: 
'#####################################################################################################################

 Function NS_EZAccess()
   	
  	Call CloseAllBrowser()
	Wait(4)
	ez_flag = True
	Systemutil.CloseProcessByName "iexplore.exe"
	SystemUtil.Run "iexplore.exe", Environment.Value("EZ_URL")
	Browser("name:=QA: Zurich Intranet Login.*").Page("title:=QA: Zurich Intranet Login.*").WebEdit("name:=username").Set Environment.Value("EZ_LoginId")
	Browser("name:=QA: Zurich Intranet Login.*").Page("title:=QA: Zurich Intranet Login.*").WebEdit("name:=password").Set Environment.Value("EZ_LoginPassword")
	Browser("name:=QA: Zurich Intranet Login.*").Page("title:=QA: Zurich Intranet Login.*").WebButton("name:=Log In").Click
	Wait(2)
	If Browser("name:=eZACCESS Start Page").Page("title:=eZACCESS Start Page").Link("name:=Cleanup Sessions").Exist Then
		Browser("name:=eZACCESS Start Page").Page("title:=eZACCESS Start Page").Link("name:=Cleanup Sessions").Click
		Wait(2)
		If Dialog("Message from webpage").WinButton("OK").Exist Then
			Dialog("Message from webpage").WinButton("OK").Click
			Wait(1)
		End If 
	End If
	Browser("name:=eZACCESS Start Page").Page("title:=eZACCESS Start Page").Link("name:=eZACCESS Production System.*").Click
	Wait(2)
	Browser("ACT II").Page("ACT II").Frame("topFrame").WebElement("eZACCESS").Click
	Wait(1)
	Browser("ACT II").Page("ACT II").Frame("topFrame").WebElement("Claim Search").Click
    Browser("ACT II").Page("ACT II").Frame("sidebarFrame").WebEdit("claimId").Set Environment.Value("ClaimNumber")
    Wait(2)
	Browser("ACT II").Page("ACT II").Frame("sidebarFrame").WebButton("Go!").Click
	Wait(2)
	If Browser("ACT II").Page("ACT II").Frame("contentFrame").WebElement("Ezaccess").Exist(2) then
		If trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebElement("Ezaccess").GetROProperty("innertext"))="Please correct the following errors before proceeding:" then
			 funarr(12) = False
			 ez_flag = False
			 Browser("ACT II").Close
			 EzRegStatus = "False"
		Else
			 funarr(12) = True
			EzRegStatus = "True"
			Browser("ACT II").Page("ACT II").Frame("sidebarFrame").Link("Claim Data").Click
			Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Claim Summary").Click
			If DataTable("ClaimsMadeInd","GL-Data") = "Yes"  Then
				Call ClaimsMadeIndicator_O
			Else
				Call ClaimsMadeIndicator_not_O
			End If			   
			Val_LossDate = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_LossDate").GetROProperty("value"))
			Val_LossDate2 = Replace(Val_LossDate,"-","")
			Val_State = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CS_State").GetROProperty("default value"))
			Val_Injury_Damage = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CS_Injury/Damage").GetROProperty("default value"))
			Val_Loss_Date1 = Replace(DataTable("CS_Accident_Date","GL-Data"),"/","")
			If instr(Val_Loss_Date1,Trim(Val_LossDate2))> 0 Then ''it used to be Val_LossDate now Val_LossDate 
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess Accdient date match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess Accdient date does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess Accdient date does not match with Pega input - Failed *" 
			End If
			If Val_State <> "" Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  State match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  State does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  State does not match with Pega input - Failed *" 
			End If
			If instr(DataTable("Inj_Description","GL-Data"),Trim(Val_Injury_Damage))> 0 Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Injury description match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Injury description does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Injury description does not match with Pega input - Failed *" 
			End If
			Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Accident").Click	
			Wait 3
			Val_Acc_Email = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_Email").GetROProperty("default value"))
			Val_Acc_Time = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_Time").GetROProperty("default value"))
			Val_Acc_ClaimType = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("Acc_ClaimType").GetROProperty("default value"))
			accDesc = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_Description").GetROProperty("value")
			accCode1 = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("Acc_Code").GetROProperty("value")
			accCode = Split(accCode1, " -")
			agentLoss = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("AgentOfLoss").GetROProperty("value")
			lossLocation = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("dtlLossLocation").GetROProperty("value")
			accDate = Split(DataTable("CS_Accident_Date","GL-Data"), "/")
			repDate = Split(Val_LossDate, "-")'' here repd
			''******************************************  This needs to be corrected************************************
			If (accDate(0) = repDate(0) And accDate(1) = repDate(1) And accDate(2) = repDate(2)) Then		
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Accident page accident date match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Accident page accident date does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess Accident page accident date does not match with Pega input - Failed *" 
			End If
			If DataTable("IN_AccDescription","GL-Data") =  Trim(accDesc) Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Accident Description match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Accident Description  does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Accident Description does not match with Pega input - Failed *" 
			End If
			If DataTable("ACC_AccCode","GL-Data") =  Trim(accCode(0)) Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Accident Code match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Accident Code does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Accident Code does not match with Pega input - Failed *" 
			End If
			If DataTable("ACC_AgentLoss","GL-Data") =  Trim(agentLoss) Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Agent of Loss match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Agent of Loss does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Agent of Loss does not match with Pega input - Failed *" 
			End If
			If instr(DataTable("ACC_LossLoc","GL-Data"),Trim(lossLocation))> 0 Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Loss Location match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Loss Location does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Loss Location does not match with Pega input - Failed *" 
			End If			
			If instr(DataTable("CO_Rep_Email","GL-Data"),Trim(Val_Acc_Email))> 0 Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Email match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Email does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Email does not match with Pega input - Failed *" 
			End If
			Acc_Time = DataTable("IN_AccidentTime1","GL-Data") & ":" &DataTable("IN_AccidentTime2","GL-Data")
			If instr(Acc_Time,Trim(Val_Acc_Time))> 0 Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Accident time match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Accident time " &Val_Acc_Time &"does not match with Pega input  "&Acc_Time
				Excel_Comments = Excel_Comments & "* Ezaccess  Accident time does not match with Pega input - Failed *"
			End If
			Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Claim References").Click
			Wait(1)
			''****Handling the pop up which appears if any change is made in any field in eZACCESS .
			If Browser("ACT II").Dialog("Message from webpage").Exist Then
				Browser("ACT II").Dialog("Message from webpage").WinButton("Cancel").Click
			End If
			Val_CR_IncidentNumber = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CR_IncidentNumber").GetROProperty("default value"))
			'If instr(Environment.Value("SCaseId"),Trim(Val_CR_IncidentNumber))> 0 Then
			If  Val_CR_IncidentNumber <> "" Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Scase ID match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Scase ID does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Scase ID does not match with Pega input - Failed *"
			End If
			Browser("ACT II").Page("ACT II").Frame("sidebarFrame").Link("Claimants").Click	
			Wait 3
			If DataTable("IN_Product","GL-Data") = "Occupational Accident"  Then  ''OCA ValId 
				ValId=Trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("ValID").GetROProperty("innertext"))
				If Instr(1,ValId,"OCMED")>0  Then
			 		ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess- Val Id for OCA Dsipalyed as " & ValId
			 	ElseIf Instr(1,ValId,"OCWGE")>0 Then
			 		ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess- Val Id for OCA Dsipalyed as " & ValId
			 	Else
			 	ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess- Wrong Val Id for OCA Dsipalyed as " & ValId
			 	End If
			End If 
			If Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Raj Ram S").Exist Then
			 Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Raj Ram S").Click	
			End If 
			Wait 3
			Val_CL_CD_Fname = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CL_CD_Fname").GetROProperty("default value"))
			Val_CL_CD_Lname = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CL_CD_Lname").GetROProperty("default value"))
			Val_CL_CD_M = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CL_CD_M").GetROProperty("default value"))
			Val_CL_CD_Name2= Val_CL_CD_Fname & Val_CL_CD_M & Val_CL_CD_Lname
			Val_CL_CD_Name1 =  DataTable("Party_Fname","GL-Data") & DataTable("Party_MI","GL-Data") & DataTable("Party_Lname","GL-Data")
			If instr(Trim(Val_CL_CD_Name1),Trim(Val_CL_CD_Name2))> 0 Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Claimant Name match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Claimant Name does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Claimant Name does not match with Pega input - Failed *" 
			End If
			If Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Injury").Exist Then
				Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Injury").Click	
				Wait 3
				Val_CL_INJ_CauseOf_Injury = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CL_INJ_CauseOf_Injury").GetROProperty("default value"))
				Val_CL_INJ_InjDes = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CL_INJ_InjDes").GetROProperty("default value"))
				Val_CL_INJ_BodyPart = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("CL_INJ_BodyPart").GetROProperty("default value"))
				Val_CL_INJ_InitialTreatment = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("CL_INJ_InitialTreatment").GetROProperty("default value"))
				Val_CL_INJ_InjCode = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("CL_INJ_InjCode").GetROProperty("default value"))
				If instr(DataTable("Inj_CauseInjury","GL-Data"),Trim(Val_CL_INJ_CauseOf_Injury))> 0 Then
					ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Cause of Injury match with Pega input"
				Else
					ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Cause of Injury does not match with Pega input"
					Excel_Comments = Excel_Comments & "* Ezaccess  Cause of Injury does not match with Pega input - Failed *"
				End If
				If instr(DataTable("Inj_Description","GL-Data"),Trim(Val_CL_INJ_InjDes))> 0 Then
					ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Injury Description match with Pega input"
				Else
					ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess   Injury Description  does not match with Pega input"
					Excel_Comments = Excel_Comments & "* Ezaccess   Injury Description  does not match with Pega input - Failed *" 
				End If
				If instr(DataTable("Inj_BodyPart","GL-Data"),Trim(Val_CL_INJ_BodyPart))> 0 Then
					ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Body Part match with Pega input"
				Else
					ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Body Part does not match with Pega input"
					Excel_Comments = Excel_Comments & "* Ezaccess  Body Part does not match with Pega input - Failed *" 
				End If
				If instr(DataTable("Inj_InitialTreatment","GL-Data"),Trim(Val_CL_INJ_InitialTreatment))> 0 Then
					ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Initial Treatment match with Pega input"
				Else
					ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Initial Treatment  does not match with Pega input"
					Excel_Comments = Excel_Comments & "* Ezaccess  Initial Treatment  does not match with Pega input - Failed *" 
				End If
				If instr(DataTable("Inj_Nature","GL-Data"),Trim(Val_CL_INJ_InjCode))> 0 Then
					ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Nature of Injury match with Pega input"
				Else
					ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Nature of Injury does not match with Pega input"
					Excel_Comments = Excel_Comments & "* Ezaccess  Nature of Injury does not match with Pega input - Failed *" 
				End If
			End If
		End if
	Else
		funarr(12) = True
		EzRegStatus = "True"
		Browser("ACT II").Page("ACT II").Frame("sidebarFrame").Link("Claim Data").Click
		If DataTable("ClaimsMadeInd","GL-Data") = "Yes"  Then
			Call ClaimsMadeIndicator_Occurrence
		Else
			Call ClaimsMadeIndicator_not_Occurrence
		End If		
		Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Claim Summary").Click
		Val_LossDate = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_LossDate").GetROProperty("value"))
		Val_LossDate2 = Replace(Val_LossDate,"-","")
		Val_State = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CS_State").GetROProperty("default value"))
		Val_Injury_Damage = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CS_Injury/Damage").GetROProperty("default value"))
		Val_Loss_Date1 = Replace(DataTable("CS_Accident_Date","GL-Data"),"/","")
		If instr(trim(Val_Loss_Date1),Trim(Val_LossDate2))> 0 Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess Accdient date match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess Accdient date does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess Accdient date does not match with Pega input - Failed *" 
		End If
		If Val_State <> "" Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  State match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  State does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess  State does not match with Pega input - Failed *" 
		End If
		If instr(DataTable("Inj_Description","GL-Data"),Trim(Val_Injury_Damage))> 0 Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Injury description match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Injury description does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess  Injury description does not match with Pega input - Failed *" 
		End If
		Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Accident").Click	
		Wait 3
		Val_Acc_Email = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_Email").GetROProperty("default value"))
		Val_Acc_Time = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_Time").GetROProperty("default value"))
		Val_Acc_ClaimType = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("Acc_ClaimType").GetROProperty("default value"))
		accDesc = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_Description").GetROProperty("value")
		accCode1 = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("Acc_Code").GetROProperty("value")
		accCode = Split(accCode1, " -")
		agentLoss = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("AgentOfLoss").GetROProperty("value")
		lossLocation = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("dtlLossLocation").GetROProperty("value")
		'''' *************************     Other Validation *********************************************************************************************************************
		accDate = Split(DataTable("CS_Accident_Date","GL-Data"), "/")
		repDate = Split(Val_LossDate, "-")
		If (accDate(0) = repDate(0) And accDate(1) = repDate(1) And accDate(2) = repDate(2)) Then		
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Accident page accident date match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Accident page accident date does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess Accident page accident date does not match with Pega input - Failed *" 
		End If
		If DataTable("IN_AccDescription","GL-Data") =  Trim(accDesc) Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Accident Description match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Accident Description  does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess  Accident Description does not match with Pega input - Failed *" 
		End If
		If DataTable("ACC_AccCode","GL-Data") =  Trim(accCode(0)) Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Accident Code match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Accident Code does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess  Accident Code does not match with Pega input - Failed *" 
		End If
		If DataTable("ACC_AgentLoss","GL-Data") =  Trim(agentLoss) Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Agent of Loss match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Agent of Loss does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess  Agent of Loss does not match with Pega input - Failed *" 
		End If
		If instr(DataTable("ACC_LossLoc","GL-Data"),Trim(lossLocation))> 0 Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Loss Location match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Loss Location does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess  Loss Location does not match with Pega input - Failed *" 
		End If			
		If instr(DataTable("CO_Rep_Email","GL-Data"),Trim(Val_Acc_Email))> 0 Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Email match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Email does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess  Email does not match with Pega input - Failed *" 
		End If
		Acc_Time = DataTable("IN_AccidentTime1","GL-Data") & ":" &DataTable("IN_AccidentTime2","GL-Data")
		If instr(Acc_Time,Trim(Val_Acc_Time))> 0 Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Accident time match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Accident time " &Val_Acc_Time &"does not match with Pega input  "&Acc_Time
			Excel_Comments = Excel_Comments & "* Ezaccess  Accident time does not match with Pega input - Failed *"
		End If
		Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Claim References").Click
		Wait 3
		Val_CR_IncidentNumber = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CR_IncidentNumber").GetROProperty("default value"))
		If  Val_CR_IncidentNumber <> "" Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Scase ID match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Scase ID does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess  Scase ID does not match with Pega input - Failed *"
		End If
		Browser("ACT II").Page("ACT II").Frame("sidebarFrame").Link("Claimants").Click	
		Wait 3
		rem Call Validate_ValID()
		If Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Raj Ram S").Exist Then
		 Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Raj Ram S").Click	
		End If 
		Wait 3
		If DataTable("IN_Product","GL-Data") = "Occupational Accident"  Then  ''OCA ValId 
			 ValId=Trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("ValID").GetROProperty("innertext"))
			 If Instr(1,ValId,"OCMED")>0  Then
			 		ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess- Val Id for OCA Dsipalyed as " & ValId
			 	ElseIf Instr(1,ValId,"OCWGE")>0 Then
			 		ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess- Val Id for OCA Dsipalyed as " & ValId
			 	Else
			 	ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess- Wrong Val Id for OCA Dsipalyed as " & ValId
			 End If
		End If 
		Val_CL_CD_Fname = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CL_CD_Fname").GetROProperty("default value"))
		Val_CL_CD_Lname = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CL_CD_Lname").GetROProperty("default value"))
		Val_CL_CD_M = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CL_CD_M").GetROProperty("default value"))
		Val_CL_CD_Name2= Val_CL_CD_Fname & Val_CL_CD_M & Val_CL_CD_Lname
		Val_CL_CD_Name1 =  DataTable("Party_Fname","GL-Data") & DataTable("Party_MI","GL-Data") & DataTable("Party_Lname","GL-Data")
		If instr(Trim(Val_CL_CD_Name1),Trim(Val_CL_CD_Name2))> 0 Then
			ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Claimant Name match with Pega input"
		Else
			ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Claimant Name does not match with Pega input"
			Excel_Comments = Excel_Comments & "* Ezaccess  Claimant Name does not match with Pega input - Failed *" 
		End If
		If Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Injury").Exist Then
			Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Injury").Click	
			Wait 3
			Val_CL_INJ_CauseOf_Injury = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CL_INJ_CauseOf_Injury").GetROProperty("default value"))
			Val_CL_INJ_InjDes = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("CL_INJ_InjDes").GetROProperty("default value"))
			Val_CL_INJ_BodyPart = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("CL_INJ_BodyPart").GetROProperty("default value"))
			Val_CL_INJ_InitialTreatment = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("CL_INJ_InitialTreatment").GetROProperty("default value"))
			Val_CL_INJ_InjCode = trim(Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("CL_INJ_InjCode").GetROProperty("default value"))
			If instr(DataTable("Inj_CauseInjury","GL-Data"),Trim(Val_CL_INJ_CauseOf_Injury))> 0 Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Cause of Injury match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Cause of Injury does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Cause of Injury does not match with Pega input - Failed *"
			End If
			If instr(DataTable("Inj_Description","GL-Data"),Trim(Val_CL_INJ_InjDes))> 0 Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Injury Description match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess   Injury Description  does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess   Injury Description  does not match with Pega input - Failed *" 
			End If
			If instr(DataTable("Inj_BodyPart","GL-Data"),Trim(Val_CL_INJ_BodyPart))> 0 Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Body Part match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Body Part does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Body Part does not match with Pega input - Failed *" 
			End If
			If instr(DataTable("Inj_InitialTreatment","GL-Data"),Trim(Val_CL_INJ_InitialTreatment))> 0 Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Initial Treatment match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Initial Treatment  does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Initial Treatment  does not match with Pega input - Failed *" 
			End If
			If instr(DataTable("Inj_Nature","GL-Data"),Trim(Val_CL_INJ_InjCode))> 0 Then
				ReportResult_Event micPass, "Invoking Business component: Ezaccess" , "Ezaccess  Nature of Injury match with Pega input"
			Else
				ReportResult_Event micFail, "Invoking Business component: Ezaccess" , "Ezaccess  Nature of Injury does not match with Pega input"
				Excel_Comments = Excel_Comments & "* Ezaccess  Nature of Injury does not match with Pega input - Failed *" 
			End If
		End If
	End If
End Function



Function Excel_ReportGeneration()

	Dim rowcount1
	Dim result_path
	Dim result_path2
	Dim Excel_Comments
	
	result_path1 =Environment.Value("RelativePath")& "\Result\Claim_Number.xls"
	Set objExcel=CreateObject("Excel.Application")
	Set objWrkbook=objExcel.Workbooks.Open(result_path1)
	Set objSheet1=objExcel.Sheets("Sheet1")
	rowcount1 = objSheet1.usedrange.rows.count
	rowcount1= rowcount1+1
	Excel_Comments = ""
	objSheet1.cells(rowcount1,1).Value=Date
	objSheet1.cells(rowcount1,2).Value=DataTable.Value("TestcaseID","BusinessFlow")
	objSheet1.cells(rowcount1,3).Value= Environment.Value("SCaseId")
	objSheet1.cells(rowcount1,4).Value= Environment.Value("ClaimNumber")
	objSheet1.cells(rowcount1,5).Value= Environment.Value("CarePA_Status")
	If EzRegStatus = "True" Then
		objSheet1.cells(rowcount1,6).Value= "PASS"
		If Claim_Number ="Number N/A" or Claim_Number = "Farmer's Policy" Then
			objSheet1.cells(rowcount1,7).Value= ""
		Else
			objSheet1.cells(rowcount1,7).Value= Environment.Value("TC_Status")
		End If	
	Else
		objSheet1.cells(rowcount1,6).Value= "FAIL"
		If Claim_Number ="Number N/A" or Claim_Number = "Farmer's Policy" Then
			objSheet1.cells(rowcount1,6).Value= ""
			objSheet1.cells(rowcount1,7).Value= ""
		Else
			'objSheet1.cells(rowcount1,6).Value= Environment.Value("TC_Status")
		End If
	End If
	objSheet1.cells(rowcount1,8).Value= Excel_Comments
	objWrkbook.Save
	objExcel.Quit
	Set objExcel=Nothing
	Set objWrkbook=Nothing
	Set objSheet1=Nothing
	
End Function

' ********************************** HealthCare Test Cases **********************************************************************
Function Re_select_Customer()

	 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Re-select Customer").Click
	 Wait 5

	 If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("CS_Search").Exist then
		  ReportResult_Event micPass, "Invoking Business component: Reselect_Customer" , "Reselect customer button is clicked and navigated to Customer search page"
	Else
			ReportResult_Event micFail, "Invoking Business component: Reselect_Customer" , "Reselect customer button is clicked and it is not navigated to Customer search page"
	End if 
	
End Function
'**********************************************************************************************************************************************************************
Function Re_select_Employee()

	 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Re-select Employee").Click
	 Wait 5
	
End Function
'**********************************************************************************************************************************************************************
Function Void_Incident()

	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Void Incident").Click
		
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Enter_VodReason").Select DataTable("Enter_VodReason","GL-Data")
		If  Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("VI_Reason").Exist Then
				Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("VI_Reason").Set DataTable("VI_Reason","GL-Data")
		End If
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("VI_Submit").Click
	Wait 3
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Confirm").Click

End Function
'**********************************************************************************************************************************************************************
Function TC18_E2E_Scenario_Void_the_incident1()

	Browser("ClaimsBrowser").Page("Inbox").Frame("topFrame").WebElement("Search_Incident_Icon").Click
	Wait 5
	 Browser("ACT II").Page("Claim CC Service Items").WebEdit("SI_Incident_Number").Set Environment.Value("SCaseId")
	 Browser("ACT II").Page("Claim CC Service Items").WebButton("SI_Search").Click
	 Wait 5
	Browser("ACT II").Page("Claim CC Service Items").WebElement("SI_Res_ScaseID").Click
	 Wait 5
	Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebButton("IS_Restore").Click

End Function
'**********************************************************************************************************************************************************************
Function TC19_E2E_Scenario_Distribution_form_validation1()

	Browser("ClaimsBrowser").Page("Inbox").Frame("topFrame").WebElement("Search_Incident_Icon").Click
	Wait 5
	 Browser("ACT II").Page("Claim CC Service Items").WebEdit("SI_Incident_Number").Set Environment.Value("SCaseId")
	 Browser("ACT II").Page("Claim CC Service Items").WebButton("SI_Search").Click
	 Wait 5
	Browser("ACT II").Page("Claim CC Service Items").WebElement("SI_Res_ScaseID").Click
	 Wait 5
	Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebButton("IS_Update Claim Data").Click

	If Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebButton("IS_Save").Exist Then
		ReportResult_Event micPass, "Invoking Business component: TC19_E2E_Scenario_Distribution_form_validation1" , "Page is navigated to 'General Information' Screen after 'Update Claim Data' button is clicked"
		Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").Link("IS_Site Details").Click
		Wait 3
		Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebEdit("IS_SiteDetails_CustName").Set "Test"
	Else
		ReportResult_Event micPass, "Invoking Business component: TC19_E2E_Scenario_Distribution_form_validation1" , "Page is not navigated to 'General Information' Screen after 'Update Claim Data' button is clicked"
	End If

	Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebButton("IS_Save").Click	
	Wait 3
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Confirm").Click
	Wait 3

	If Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebList("none").Exist Then
		ReportResult_Event micPass, "Invoking Business component: TC19_E2E_Scenario_Distribution_form_validation1" , "Page is navigated to Inbox after Confirm button is clicked"
	Else
		ReportResult_Event micFail, "Invoking Business component: TC19_E2E_Scenario_Distribution_form_validation1" , "Page is not navigated to Inbox after Confirm button is clicked"
	End If

End Function
'**********************************************************************************************************************************************************************
Function TC25_Close_and_Reselect_Customer_Property_Damage1()

	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("ACC_Next>>").Click
	Wait 3
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Close").Click
	Wait 3

	If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist then
		Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("Cancel").Click
		Wait 3
		ReportResult_Event micPass, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "PopUp is present when Close button is clicked"
	Else
		ReportResult_Event micFail, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "PopUp is not present when Close button is clicked"
	End if

	If Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebList("none").Exist Then
		ReportResult_Event micPass, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "Page is navigated to Inbox after Cancel button is clicked from the Popup"
		Status = trim(Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebTable("Urgency").GetCellData(2,2))
		ScaseID = trim(Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebTable("Urgency").GetCellData(2,7))
		
			If Status = "Pending" and ScaseID = Environment.Value("SCaseId") Then
				Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebElement("IB_IncidentID").Click
				ReportResult_Event micPass, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "WorkItem with the Status = 'Pending' and IncidentID ="& Environment.Value("SCaseId") & "is present in the Inbox page"
					If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Re-select Customer").Exist Then
						ReportResult_Event micPass, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "Page is navigated to PropertyDamage after selecting the pending workitem from the Inbox page"
					Else
						ReportResult_Event micFail, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "Page is not navigated to PropertyDamage after selecting the pending workitem from the Inbox page"
					End If
			Else
				ReportResult_Event micFail, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "WorkItem with the Status = 'Pending' and IncidentID ="& Environment.Value("SCaseId") & "is not present in the Inbox page"
			End If

	Else
		ReportResult_Event micFail, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "Page is not navigated to Inbox after Cancel button is clicked from the Popup"
	End If

End Function
'**********************************************************************************************************************************************************************
Function TC02_E2E_scenario()

	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").Link("Incident").Click
	Wait 4
	counter = Environment.Value("counter")
	counter = counter + 1
	DataTable.GetSheet("GL-Data").SetCurrentRow(counter)

End Function
'**********************************************************************************************************************************************************************
Function Val_ClaimSeries_9()

	Claim_Series = Left(Environment.Value("ClaimNumber"),1)

	If Claim_Series = "9" Then 
			ReportResult_Event micPass, "Invoking Business component: Assignment" , "System generates Claim Series = 9 Subpath = 9"
	Else
			ReportResult_Event micFail, "Invoking Business component: Assignment" , "System did not generates Claim Series = 9 Subpath = 9"
	End If

End Function
'**********************************************************************************************************************************************************************
Function Val_ClaimSeries_6()

	Claim_Series = Left(Environment.Value("ClaimNumber"),1)

	If Claim_Series = "6" Then 
			ReportResult_Event micPass, "Invoking Business component: Assignment" , "System generates Claim Series = 6 Subpath = 65"
	Else
			ReportResult_Event micFail, "Invoking Business component: Assignment" , "System did not generates Claim Series = 6 Subpath = 65"
	End If

End Function

Function TPA_Override()

		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Click
		Wait(1)

		If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("CO_Rep_Name").Exist then
				ReportResult_Event micPass, "Invoking Business component: TPA_Override" , "TPA Override button is clicked and navigated to Contact Info page"
		Else
				ReportResult_Event micFail, "Invoking Business component: TPA_Override" , "TPA Override button is clicked and it is not navigated to Contact Info page"
		End if 


End Function
'**********************************************************************************************************************************************************************

Function Policy_Override()

	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("Policy_Override").Exist then
		ReportResult_Event micPass, "Invoking Business component: Policy_Override" , "Invoking Business component: Policy_Override - Done"
	Else
			ReportResult_Event micFail, "Invoking Business component: Policy_Override" , "Invoking Business component: Policy_Override - Failed"
			Excel_Comments = Excel_Comments & "* Invoking Business component: Policy_Override - Failed *" 
	End if	

End Function
'********************************************************************************************************************************************************************************
Function Void_Reason()

		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Void Incident").Click
		Wait 3
		If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Void_Reason").Exist then
			Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Void_Reason").Select "Duplicate Claim"
			Wait(1)
		End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("VR_Reason").Set "test"
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("VR_Submit").Click
		Wait(1)

		If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Confirm").Exist then
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Confirm").Click
				Wait(1)
				ReportResult_Event micPass, "Invoking Business component: Void_Reason" , "Invoking Business component: Void_Reason - Done"
		Else
				ReportResult_Event micFail, "Invoking Business component: Void_Reason" , "Invoking Business component: Void_Reason - Failed"
				Excel_Comments = Excel_Comments & "* Invoking Business component: Void_Reason - Failed *" 
		End if	

End Function
'*******************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************
Function Set_Diaction_to_RoomPane()

	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").SetTOProperty "name","RoomPane"

End Function
'********************************************************************************************************************************************************************************