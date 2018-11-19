'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
								'Carepoint WC LOB Business Functions
								'Created By : Srirekha Talasila
								'Created On : 12/05/2016
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'Dim blnverify, Item_count, i, Accident_SiteAddr, pol_flag, Claim_Number
'Dim flag_arr(100),flag_EZAccPolicyNum, Excel_Comments, PegaPolicyNum, EmployerName,ChannelFlag, Acc_AccidentState,TPA_override,EzRegStatus

TPA_override = false

Function Login()

	Environment.value("str_ScreenName") = "Carepoint >>>> Login Page "
	Systemutil.CloseProcessByName "iexplore.exe"
	Systemutil.Run "iexplore.exe", Environment.Value("CP_URL")
	Set objLoginPage = Browser("ClaimsBrowser").Page("LoginPage")
	objLoginPage.WebEdit("UserIdentifier").Set Environment.Value("CP_LoginId")
	objLoginPage.WebEdit("Password").Set Environment.Value("CP_LoginPassword") 
	objLoginPage.WebButton("Log In").Click

End Function

Function Select_WorkersCompensation()

	Environment.value("str_ScreenName") = "Carepoint >>>> Select LOB"
	Set obj_BrowserPage = Browser("ClaimsBrowser").Page("Inbox")
	obj_BrowserPage.Link("My Group").Click
	obj_BrowserPage.WebList("select").Select "Workers Compensation"
	
End function


Function Select_Workitem()

	Environment.value("str_ScreenName") = "Carepoint >>>> WC - Select WorkItem "
	Wait(3)
'	Browser("Customer_Browser").Page("WorkList_Basket").WebElement("innertext:=Click to sortDate/Time received.*").Click
	Browser("Customer_Browser").Page("WorkList_Basket").WebElement("title:=Click.*","Index:=12").click
	wait(6)
	SelectionCount=1
	Do
         If SelectionCount=1 Then
				Set tabobj=Browser("Customer_Browser").Page("WorkList_Basket").WebTable("WorkItem_Selection")
				rowcount=trim(tabobj.GetROProperty("rows"))
				For row=2 To rowcount 
					Set tabobj=Browser("Customer_Browser").Page("WorkList_Basket").WebTable("WorkItem_Selection")
					Status=Trim(tabobj.GetCellData(row,3))
					Channel=Trim(tabobj.GetCellData(row,4))
					IncidentID=Trim(tabobj.GetCellData(row,8))	
					IDType=left(IncidentID,1)
					currentrowcount=row
					If Status="New"  and IDType<>"S"  and Channel <> "WEB"  and Channel<>  "FTP" Then	
						Set objref=createobject("Mercury.DeviceReplay")
						x=Browser("Customer_Browser").Page("WorkList_Basket").WebTable("WorkItem_Selection").ChildItem(row,3,"WebElement",0).GetRoProperty("abs_x")
			 			y=Browser("Customer_Browser").Page("WorkList_Basket").WebTable("WorkItem_Selection").ChildItem(row,3,"WebElement",0).GetRoProperty("abs_y")
			 			objref.MouseDblClick x,y,0 
			 			wait(3)
						objref.MouseDblClick x,y,0 			 			
			 			Set objref=nothing		
						Exit For				 			
	         		End If	
					If row=13 Then
				   		 Set obj = CreateObject("WScript.Shell")
				  		 obj.SendKeys ("{PGDN}")
				  		 Set obj=nothing 	
					End If
					If  row=Cint(rowcount) Then
						CustomerSearchCheck=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Customer Search").Exist(5)
						If  CustomerSearchCheck=False Then
							Browser("Customer_Browser").Page("WorkList_Basket").Link("Next").Click
							row=1
						End If
					End If
				Next
				Set tabobj=nothing
				SelectionCount=SelectionCount+1	
		Else
			rem SelectionCount=SelectionCount+1
         End If
		
			Check=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Customer Search").Exist(5)
			 If Check="True" Then
					Exit Do
			End If
	Loop Until Check=False	
	
	' Clicking the Customer Search Button
	 
	If Browser("title:=TestDaeja.*").Exist(3) Then
		Browser("title:=TestDaeja.*").Close 
	    Wait(1)
	End If
	If DataTable("Customer_Employee_SearchFlag","CommonTestData") = "TRUE" Then
		Customer_Search()
	Else
		Employee_Search()
	End If
	
End Function


Function Customer_Search()

	Environment.value("str_ScreenName") = "Carepoint >>>> WC - Customer Search "
	Dim objBrwpage_CustomerSearch

	set objBrwpage_CustomerSearch=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame") 
	objBrwpage_CustomerSearch.WebButton("Customer Search").Click 

	If (DataTable("Add_NewCustomer_Flag","CommonTestData") = "FALSE") Then
		Browser("ClaimsBrowser").Page("Inbox").Sync
		If objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Exist(5) Then
		   objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Set DataTable("CS_Accident_Date","CommonTestData")	
		End If		
		objBrwpage_CustomerSearch.WebEdit("CS_CustomerName").Set DataTable("CS_Customer_Name","CommonTestData")
		objBrwpage_CustomerSearch.WebEdit("CS_SiteCode").Set DataTable("CS_SiteCode","CommonTestData")
'		objBrwpage_CustomerSearch.WebEdit("CS_PolicyNo").Set DataTable("CS_PolicyNo","CommonTestData")
		objBrwpage_CustomerSearch.WebButton("CS_Search").Click		
		Wait(5)
		Index=1
		while index<>0'''Here the condition will waits till Web Table load
			If (objBrwpage_CustomerSearch.webelement("CS_No_Matching_Data").Exist(30) ) Then
				index=0
			Else
				index=0
				Set obj_BusinessUnit=Browser("CreationTime:=0").Page("title:=.*").Frame("name:=actionIFrame").WebTable("column names:=Click to sortBusiness Unit ,;Click to sortCustomer Name ,;Click to sortEntity Name ,;Click to sortSite Name ,;Click to sortSite Code ,;Click to sortAddress 1 ,;Click to sortAddress 2 ,;Click to sortCity ,;Click to sortState ,;Click to sortZip Code ,;Click to sortPhone ,;Click to sortFax ,","index:=23").ChildItem(2,1,"WebElement",0)''@DP
				If obj_BusinessUnit.Exist(30) Then
					obj_BusinessUnit.click '''This will target first row in the Customer SEarch result 
				End If
				wait(3)
				If objBrwpage_CustomerSearch.WebButton("html id:=submitButton").Exist(30) Then
					Setting.WebPackage("ReplayType") = 1
					objBrwpage_CustomerSearch.WebButton("html id:=submitButton").Click
					Setting.WebPackage("ReplayType") = 2
				End If
				If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(5) then
					Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click				
				End if
			End If 	
		Wend
		
		If Browser("title:=Care.*").Exist(5) Then
			   Browser("title:=Care.*").Close 
			   Wait(1)
	   	End If  	
		If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Click
		End If
		Wait(2)
	ElseIf DataTable("Add_NewCustomer_Flag","CommonTestData") = "TRUE" Then		
		Add_NewCustomer()
	Else
		'Do Nothing
	End If
	
End Function

Function Employee_Search()

		Dim obj_actionIFrame
		Set obj_InboxPage = Browser("ClaimsBrowser").Page("Inbox")
		Set obj_actionIFrame = obj_InboxPage.Frame("actionIFrame")
		obj_actionIFrame.WebButton("Employee_Search").Click
		obj_actionIFrame.WebList("ES_CustomerName").Select DataTable("ES_CustomerName","CommonTestData")
        obj_actionIFrame.WebList("ES_ClaimType").Select DataTable("ES_ClaimType","CommonTestData")
        obj_actionIFrame.WebList("ES_ReportingType").Select DataTable("ES_ReportingType","CommonTestData")
        obj_actionIFrame.WebEdit("ES_Emp_LastName").Set DataTable("ES_Emp_LastName","CommonTestData")
        obj_actionIFrame.WebEdit("ES_EmpID").Set DataTable("ES_EmpID","CommonTestData")
        obj_actionIFrame.WebButton("ES_Search").Click
		cell_data1 = obj_actionIFrame.WebTable("ES_SearchResults").GetCellData(2,2)
		If cell_data1 <> "" Then
			Set empobj = obj_actionIFrame.WebTable("ES_SearchResults")
			Set empobj2 = empobj.ChildItem(2,1,"WebRadioGroup",0)				
			class_name = empobj2.getroproperty("class")
			If class_name = "Radio lvInputSelection" Then
				obj_actionIFrame.WebRadioGroup("ES_RadioButton").Click
			Else
				 'Do Nothing
			End if
		End If
		obj_actionIFrame.WebButton("ES_Select").Click
		If ChannelFlag <> "FTP" Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("actionIFrame").WebButton("html id:=startProcessButton","title:=Complete this assignment","name:=.*Start Process.*").Click
		End If

End Function

Function Add_NewCustomer()

		Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Add New Customer "
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Add_New_Customer").Click
        Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("Addcustomer_CustomerName").Set DataTable("AddCustomer_CustomerName","CommonTestData")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCustomer_Addr1").Set DataTable("AddCustomer_Addr1","CommonTestData")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCustomer_ZIP").Set DataTable("AddCustomer_ZIP","CommonTestData")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCustomer_Phone").Set DataTable("AddCustomer_Phone","CommonTestData")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCustomer_Email").Set DataTable("AddCustomer_Email","CommonTestData")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("html id:=taxId").Set DataTable("AddCustomer_EmpTaxID","CommonTestData")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCustomer_SiteCode").Set DataTable("AddCustomer_SiteCode","CommonTestData")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("AddCustomer_Submit").Click
		wait(2)
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("html id:=startProcessButton","title:=Complete this assignment","name:=.*Start Process.*").Click
		
End Function

Function Incident()

	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Incident Screen "

	Set obj_InboxPage = Browser("ClaimsBrowser").Page("Inbox")
	Set obj_DiactionFrame = obj_InboxPage.Frame("DIACTION")
	If len(Trim(obj_DiactionFrame.WebEdit("Site_TIN").GetROProperty("value"))) =4 Then 
		obj_DiactionFrame.WebEdit("Site_TIN").Set DataTable("TIN_Number","CommonTestData")
	End If

	'Reporter Information
	obj_DiactionFrame.WebEdit("CO_Rep_Name").Set DataTable("CO_Rep_Name","CommonTestData")
	obj_DiactionFrame.WebEdit("CO_Rep_Email").Set DataTable("CO_Rep_Email","CommonTestData")
	obj_DiactionFrame.WebEdit("CO_Rep_Phone").Set DataTable("CO_Rep_Phone","CommonTestData")
	obj_DiactionFrame.WebList("Contact_RelationshipToClaim").Select DataTable("CO_Rep_Relationship","CommonTestData")
	'Customer Contact Information
	obj_DiactionFrame.WebEdit("Customer_Contact_Name").Set DataTable("Customer_Contact_Name","CommonTestData")
	obj_DiactionFrame.WebEdit("CO_CusCon_Email").Set DataTable("CO_CusCon_Email","CommonTestData")
	obj_DiactionFrame.WebEdit("CO_CusCon_Fax").Set DataTable("CO_CusCon_Fax","CommonTestData")
	obj_DiactionFrame.WebEdit("CO_CusCon_Phone").Set DataTable("CO_CusCon_Phone","CommonTestData")
	'Incident Details
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pIncidentDetails.*pDateOfLoss").Set "10/10/1989 "
	If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(3) then
		Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click
	End If
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pIncidentDetails.*pDateOfLoss").Set DataTable("CS_Accident_Date","CommonTestData")
	obj_DiactionFrame.WebList("AccidentState").Select DataTable("AccidentState","CommonTestData")
	If DataTable("LetRest","CommonTestData") = "TRUE" Then
		obj_DiactionFrame.WebCheckBox("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pIncidentDetails.*pReportOnly","type:=checkbox").Set "ON"
	End If
	obj_DiactionFrame.WebList("AccidentTime1").Select DataTable("AccidentTime1","CommonTestData")
	obj_DiactionFrame.WebList("AccidentTime2").Select DataTable("AccidentTime2","CommonTestData")
	obj_DiactionFrame.WebList("AccidentTime3").Select DataTable("AccidentTime3","CommonTestData")
	obj_DiactionFrame.WebEdit("AccidentDescription").Set DataTable("AccDescription","CommonTestData")
	obj_DiactionFrame.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pIncidentDetails.*pclaimSubType").Select DataTable("ClaimType","CommonTestData")
	 
	'Claimant Details
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pFirstName").Set DataTable("FirstName","CommonTestData")
	obj_DiactionFrame.WebEdit("LastName").Set DataTable("LastName","CommonTestData")
	obj_DiactionFrame.WebEdit("html id:=MiddleName").Set DataTable("IN_MI","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pAddressDetails.*paddressLine1").Set DataTable("CO_Claimant_Address1","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pAddressDetails.*paddressLine2").Set DataTable("CO_Claimant_Address2","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pAddressDetails.*ppostalCode").Set DataTable("CO_Claimant_Zip","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pTelNbr.*gPrimaryPhone.*pPhone").Set DataTable("CO_Claimant_HomePhone","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pTelNbr.*gAlternatePhone.*pPhone").Set DataTable("CO_Claimant_WorkPhone","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pTelNbr.*gFax.*pFax").Set DataTable("CO_Claimant_Fax","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pCellPhone").Set DataTable("CO_Claimant_CellPhone","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pEmailAddr").Set DataTable("CO_Claimant_Email","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pDateOfBirth").Set DataTable("CO_Calimant_DOB","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pSSN").Set DataTable("CO_Claimant_SSN","CommonTestData")
	obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pNumberOfDependents").Set DataTable("CO_Claimant_Dependent","CommonTestData")
	obj_DiactionFrame.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pGender").Select DataTable("CO_Claimant_Gender","CommonTestData")
	obj_DiactionFrame.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pMaritalStatus").Select DataTable("CO_Claimant_Marital","CommonTestData")
	
	x=obj_DiactionFrame.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pDistributionPreference").getroproperty("abs_x")   
	y=obj_DiactionFrame.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pDistributionPreference").getroproperty("abs_y")
	Set objref = createobject("Mercury.DeviceReplay")
	obj_DiactionFrame.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pDistributionPreference").Click
	objref.MouseClick x,y,0
	obj_DiactionFrame.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pContactInfo.*pDistributionPreference").Select DataTable("CO_Claimant_Distribution","CommonTestData")
	Set objref = nothing

	If obj_DiactionFrame.WebEdit("Emp_ZIP").Exist(3) then 
		obj_DiactionFrame.WebEdit("Emp_ZIP").Set 12345
		Set WshShell = CreateObject("WScript.Shell")
		Wait(1)
		WshShell.SendKeys "{TAB}"
		Wait(2)
		Set WshShell = Nothing
	End If 
	obj_DiactionFrame.WebButton("Ass_Save").Click
	 
	obj_DiactionFrame.WebButton("Next>>").Click
	'If Duplicate Claim Exists
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist(4) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Click
	Else 
		'Do Nothing
	End If
	

End Function


Function PolicySearch()

	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Policy Search Screen "
	Set obj_DiactionFrame = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	If DataTable("CS_PolicyNumber","CommonTestData") <> ""  Then
		obj_DiactionFrame.WebEdit("Policy_PolicyNo").Set DataTable("CS_PolicyNumber","CommonTestData")
		If obj_DiactionFrame.WebButton("Policy_Retrieve").GetROProperty ("disabled") = 0 Then
			obj_DiactionFrame.WebButton("Policy_Retrieve").Click
		End If			
	End If
	If obj_DiactionFrame.WebElement("innertext:=No matching policy records found.*","innerhtml:=No matching policy records found.*").Exist(15) Then
		pol_flag = True				
		obj_DiactionFrame.WebRadioGroup("Policy_RadioButton").Select "Indeterminate"
		PegaPolicyNum="Indeterminate"
	Else	
		cell_data = obj_DiactionFrame.WebTable("Policy_Table").GetCellData(2,1)
		PegaPolicyNum = Trim(obj_DiactionFrame.WebTable("Policy_Table").GetCellData(2,2))		
		If cell_data = "" Then
			Set polobj = obj_DiactionFrame.WebTable("Policy_Table")
			Set polobj2 = polobj.ChildItem(2,1,"WebRadioGroup",0)					
			d = polobj2.getroproperty("class")		
			If d = "Radio lvInputSelection" Then
			obj_DiactionFrame.WebRadioGroup("Policy_RadioButton").Click
			Else
				'Do Nothing
			End If
		End If
	End If 
	obj_DiactionFrame.WebButton("Next>>").Click
	

End Function



Function Override_TPA()
	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Override TPA "
	Set Obj_TPAButton = Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:= Override TPA","innertext:= Override TPA")
	If Obj_TPAButton.Exist(5) then
		Obj_TPAButton.Click
	Else
		'Do Nothing
	End If
	
End Function
	

Function Accident_Page()
	
	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Accident Screen "

	Set obj_DiactionFrame = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	obj_DiactionFrame.WebButton("Ass_Save").Click
	
	If (DataTable("Monopolistic_Override","CommonTestData") = "TRUE") Then
		obj_DiactionFrame.WebCheckBox("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAccident.*pMonopolisticOverride").Set "ON"
	End If
	
	x=obj_DiactionFrame.WebList("ACC_Fatality").getroproperty("abs_x")
	y=obj_DiactionFrame.WebList("ACC_Fatality").getroproperty("abs_y")
	Set objref = createobject("Mercury.DeviceReplay")
	obj_DiactionFrame.WebList("ACC_Fatality").click
	objref.MouseClick x,y,0
	obj_DiactionFrame.WebList("ACC_Fatality").Select DataTable("ACC_Fatality","CommonTestData")
	Set objref = nothing
	
	x=obj_DiactionFrame.WebList("ACC_AccCode").getroproperty("abs_x")
	y=obj_DiactionFrame.WebList("ACC_AccCode").getroproperty("abs_y")
	Set objref = createobject("Mercury.DeviceReplay")
	obj_DiactionFrame.WebList("ACC_AccCode").click
	objref.MouseClick x,y,0
	obj_DiactionFrame.WebList("ACC_AccCode").Select DataTable("ACC_AccCode","CommonTestData")
	Set objref = nothing
	
	x=obj_DiactionFrame.WebList("ACC_AgentLoss").getroproperty("abs_x")
	y=obj_DiactionFrame.WebList("ACC_AgentLoss").getroproperty("abs_y")
	Set objref = createobject("Mercury.DeviceReplay")
	obj_DiactionFrame.WebList("ACC_AgentLoss").click
	objref.MouseClick x,y,0
	obj_DiactionFrame.WebList("ACC_AgentLoss").Select DataTable("ACC_AgentLoss","CommonTestData")
	Set objref = nothing
	
	x=obj_DiactionFrame.WebList("ACC_LossLoc").getroproperty("abs_x")
	y=obj_DiactionFrame.WebList("ACC_LossLoc").getroproperty("abs_y")
	Set objref = createobject("Mercury.DeviceReplay")
	obj_DiactionFrame.WebList("ACC_LossLoc").click
	objref.MouseClick x,y,0
	obj_DiactionFrame.WebList("ACC_LossLoc").Select DataTable("ACC_LossLoc","CommonTestData")
	Set objref = nothing
	obj_DiactionFrame.WebList("ACC_BenefitState").Select DataTable("ACC_BenefitState","CommonTestData")
	obj_DiactionFrame.WebList("ACC_SiteAddress").Select DataTable("ACC_SiteAddress","CommonTestData")
	Accident_SiteAddr = DataTable("ACC_SiteAddress","CommonTestData")
	If  (Accident_SiteAddr = "No") Then
		obj_DiactionFrame.WebEdit("ACC_AccAddress1").Set DataTable("ACC_AccAddress1","CommonTestData")
		obj_DiactionFrame.WebEdit("ACC_AccAddress2").Set DataTable("ACC_AccAddress2","CommonTestData")
		obj_DiactionFrame.WebEdit("ACC_AccZip").Set DataTable("ACC_AccZip","CommonTestData")
		Acc_AccidentState = obj_DiactionFrame.WebList("ACC_State").GetROProperty ("value")
	Else
		Acc_AccidentState = obj_DiactionFrame.WebList("ACC_State").GetROProperty ("value")
	End If
	obj_DiactionFrame.WebEdit("ACC_Comments").Click 
	obj_DiactionFrame.WebEdit("ACC_Comments").Set DataTable("ACC_Comments","CommonTestData")
	' POLICE
	If DataTable("ACC_Police","CommonTestData")="ON"  Then
		obj_DiactionFrame.WebCheckBox("ACC_Police").Click
	End If
	obj_DiactionFrame.WebCheckBox("ACC_Fire").Set DataTable("ACC_Fire","CommonTestData")
	obj_DiactionFrame.WebCheckBox("ACC_Ambulance").Set DataTable("ACC_Ambulance","CommonTestData")
	obj_DiactionFrame.WebCheckBox("ACC_Other").Set DataTable("ACC_Other","CommonTestData")
	If DataTable("ACC_Police","CommonTestData") = "ON" Then
		obj_DiactionFrame.WebEdit("ACC_Pol_AuthName").Set DataTable("ACC_Pol_AuthName","CommonTestData")
		obj_DiactionFrame.WebEdit("ACC_Pol_OffName").Set DataTable("ACC_Pol_OffName","CommonTestData")
		obj_DiactionFrame.WebEdit("ACC_Pol_OffBatch").Set DataTable("ACC_Pol_OffBatch","CommonTestData")
		obj_DiactionFrame.WebEdit("ACC_Pol_Report").Set DataTable("ACC_Pol_Report","CommonTestData")
		obj_DiactionFrame.WebEdit("ACC_Pol_OffPhone").Set DataTable("ACC_Pol_OffPhone","CommonTestData")
		obj_DiactionFrame.WebEdit("ACC_Pol_NCIC").Set DataTable("ACC_Pol_NCIC","CommonTestData")
	ElseIf ((DataTable("ACC_Fire","CommonTestData") = "ON") OR (DataTable("ACC_Ambulance","CommonTestData") = "ON") OR (DataTable("ACC_Other","CommonTestData") = "ON")) Then
		obj_DiactionFrame.WebEdit("ACC_Ambu_AuthName").Set DataTable("ACC_Ambu_AuthName","CommonTestData")
		obj_DiactionFrame.WebEdit("ACC_Ambu_Report").Set DataTable("ACC_Ambu_Report","CommonTestData")
		obj_DiactionFrame.WebEdit("ACC_Ambu_OSHA").Set DataTable("ACC_Ambu_OSHA","CommonTestData")
	End If
	 
	obj_DiactionFrame.WebButton("Next>>").Click 

End function


Function Employment()
	
	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Employment Screen "

	Set obj_DiactionFrame = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	obj_DiactionFrame.WebList("Emp_Who").Select DataTable("EMP_Employer_Same_as_Site","CommonTestData")
	If(DataTable("EMP_Employer_Same_as_Site","CommonTestData") = "Other") Then
		
		obj_DiactionFrame.WebEdit("Emp_Name").Set DataTable("EMP_Employer_Name","CommonTestData")
		
		obj_DiactionFrame.WebEdit("Emp_Addr1").Set DataTable("EMP_Employer_Add1","CommonTestData")
		obj_DiactionFrame.WebEdit("Emp_Addr2").Set DataTable("EMP_Employer_Add2","CommonTestData")
		
		If obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployerAddress.*ppostalCode").Exist(3) then 
			obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployerAddress.*ppostalCode").Set DataTable("CS_ZipCode","CommonTestData")
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.SendKeys "{TAB}"
			Set WshShell = Nothing
		End If
		obj_DiactionFrame.WebEdit("Emp_Phone").Set DataTable("EMP_Employer_Phone","CommonTestData")
		obj_DiactionFrame.WebEdit("Emp_TIN").Set DataTable("EMP_Employer_TIN","CommonTestData")
	End If
	
	EmployerName = obj_DiactionFrame.WebEdit("Emp_Name").GetROProperty ("value") 
	If Len(EmployerName) > 40 Then
		EmployerName = Left(EmployerName,10)
		obj_DiactionFrame.WebEdit("Emp_Name").Set EmployerName
	End If
	obj_DiactionFrame.WebEdit("Emp_Reg_Occupation").Set DataTable("Employee_RegularOccupation","CommonTestData")
	obj_DiactionFrame.WebEdit("Emp_Inj_Occupation").Set DataTable("Emp_Injury_Occupation","CommonTestData")
	obj_DiactionFrame.WebEdit("Emp_ID").Set DataTable("Employee_ID","CommonTestData")
	
	If DataTable("Occupation_Category","CommonTestData") = "Yes"Then
		obj_DiactionFrame.WebButton("Ass_Save").Click
		obj_DiactionFrame.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pOccupationCatDesc").Select DataTable("Selection1","CommonTestData")
	End If
	
	obj_DiactionFrame.WebEdit("Emp_Dept").Set DataTable("Employee_Dept","CommonTestData")
	obj_DiactionFrame.WebEdit("Emp_HireDate").Set DataTable("Employee_HireDate","CommonTestData")
	obj_DiactionFrame.WebList("Emp_Status").Select DataTable("Employee_Status","CommonTestData")
	obj_DiactionFrame.WebEdit("Emp_SupervisorName").Set DataTable("Employee_SupervisorName","CommonTestData")
	obj_DiactionFrame.WebEdit("Emp_SupervisorPhone").Set DataTable("Employee_SupervisorPhone","CommonTestData")
	obj_DiactionFrame.WebEdit("Emp_NotifiedDate").Set DataTable("Employee_NotifiedDate","CommonTestData")
	obj_DiactionFrame.WebList("Emp_LostTime").Select DataTable("Employee_LostTime","CommonTestData")
	obj_DiactionFrame.WebEdit("Emp_WageAmount").Set DataTable("Employee_WageAmount","CommonTestData")
	obj_DiactionFrame.WebList("Emp_Hourly").Select DataTable("Employee_Hourly","CommonTestData")
	obj_DiactionFrame.WebEdit("Emp_Hours").Set DataTable("Employee_Hours","CommonTestData")
	obj_DiactionFrame.WebList("Emp_Days").Select DataTable("Employee_Days","CommonTestData")
	obj_DiactionFrame.WebList("Emp_WorkShift").Select DataTable("Employee_WorkShift","CommonTestData")
	If obj_DiactionFrame.WebElement("innertext:=Disability Information ","class:=subheaderLegendStyle","html tag:=LEGEND").Exist(5) then
		obj_DiactionFrame.WebEdit("Emp_LDW").Set DataTable("Employee_LDW","CommonTestData")
		obj_DiactionFrame.WebEdit("Emp_DisabilityDate").Set DataTable("Employee_DisabilityDate","CommonTestData")
		obj_DiactionFrame.WebEdit("Emp_PaidThrough_Date").Set DataTable("Employee_PaidDate","CommonTestData")
		If obj_DiactionFrame.WebList("Emp_RTW_Ind").Exist(5)  Then
			If (DataTable("ACC_Fatality","CommonTestData") = "Unknown" AND  DataTable("Employee_LostTime","CommonTestData") = "Unknown") Then
				'Do Nothing
			Else
				obj_DiactionFrame.WebList("Emp_RTW_Ind").Select DataTable("Employee_RTW_Ind","CommonTestData")
				If (DataTable("Employee_RTW_Ind","CommonTestData") = "Yes") Then
					obj_DiactionFrame.WebEdit("Emp_RTW_Date").Set DataTable("Employee_RTW_Date","CommonTestData")
				ElseIf  (DataTable("Employee_RTW_Ind","CommonTestData") = "No") Then
					obj_DiactionFrame.WebEdit("Emp_Est_RTW_Date").Set DataTable("Employee_Est_RTW_Date","CommonTestData")
				Else
					'Do Nothing
				End If 
				If obj_DiactionFrame.WebList("Emp_RTW_Qualifier").Exist(5) Then
					obj_DiactionFrame.WebList("Emp_RTW_Qualifier").Select DataTable("Employee_RTW_Qualifier","CommonTestData")	
				End If
				obj_DiactionFrame.WebEdit("Emp_ReleaseWorkDate").Set DataTable("Employee_ReleaseDate","CommonTestData")
			End If
		End If
	End If
	 
	obj_DiactionFrame.WebButton("Next>>").Click
	If  not obj_DiactionFrame.WebEdit("Inj_Description").Exist(3)Then 
		Dim i,x,obj,oDesc
		Set oDesc = Description.Create
		oDesc("micclass").value = "WebButton"
		Set obj = obj_DiactionFrame.ChildObjects(oDesc)
		For i = 0 to obj.Count - 1	 
			x = obj(i).GetROProperty("innertext") 
			If x="Next >>" Then
				obj(i).click
				Exit For 
			End If			
		Next
		Set obj =Nothing 
		Set oDesc = Nothing			
	End If
	
End Function


Function Injury()
 
 	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Injury Screen "
	
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("NatureofInjury").Select "#10"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("BodyPart").Select "#2"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Injury_Description").Set	"Entered by Automation"
	
	Set obj_DiactionFrame = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	obj_DiactionFrame.WebEdit("Inj_Description").Set DataTable("Injury_AccidentDescription","CommonTestData")
	If obj_DiactionFrame.WebElement("class:=dataValueRead","innertext:=Yes","innerhtml:=Yes").Exist(5) then
		obj_DiactionFrame.WebEdit("Inj_DeathDate").Set DataTable("Injury_DeathDate","CommonTestData")
	End if
	obj_DiactionFrame.WebList("Inj_InitialTreatment").Select DataTable("Injury_InitialTreatment","CommonTestData")
	obj_DiactionFrame.WebList("Inj_Surgery").Select DataTable("Injury_Surgery","CommonTestData")
	obj_DiactionFrame.WebList("Inj_Previous").Select DataTable("Injury_Previous","CommonTestData")
	obj_DiactionFrame.WebList("Inj_KnownMedicine").Select DataTable("Injury_KnownMedicine","CommonTestData")
	obj_DiactionFrame.WebList("Inj_Evacuation").Select DataTable("Injury_Evacuation","CommonTestData")
	obj_DiactionFrame.WebList("Inj_DrugProgram").Select DataTable("Injury_DrugProgram","CommonTestData")
	If (DataTable("Injury_DrugProgram","CommonTestData") = "Yes") Then
		obj_DiactionFrame.WebList("Inj_DrugResultPositive").Select DataTable("Injury_DrugResultPositive","CommonTestData")
	End If
	If obj_DiactionFrame.WebElement("Class Name:=titleBarLabelStyleExpanded","innertext:=Physician","html tag:=SPAN").Exist(5) then
		obj_DiactionFrame.WebEdit("Inj_Physician_FirstName").Set DataTable("Injury_Physician_FirstName","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Physician_LastName").Set DataTable("Injury_Physician_LastName","CommonTestData")
		obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pInjury.*pPhysician.*pAddr.*paddressLines.*l1").Set DataTable("Injury_Physician_Addr1","CommonTestData")
		obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pInjury.*pPhysician.*pAddr.*paddressLines.*l2").Set DataTable("Injury_Physician_Addr2","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Physician_City").Set DataTable("Injury_Physician_City","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Physician_Zip").Set DataTable("Injury_Physician_Zip","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Physician_Phone").Set DataTable("Injury_Physician_Phone","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Physician_Email").Set DataTable("Injury_Physician_Email","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Physician_Fax").Set DataTable("Injury_Physician_Fax","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Physician_MI").Set DataTable("Injury_Physician_MI","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Hospital_Name").Set DataTable("Injury_Hospital_Name","CommonTestData")
		obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pInjury.*pHospital.*pAddr.*paddressLines.*l1").Set DataTable("Injury_Hospital_Addr1","CommonTestData")
		obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pInjury.*pHospital.*pAddr.*paddressLines.*l2").Set DataTable("Injury_Hospital_Addr2","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Hospital_City").Set DataTable("Injury_Hospital_City","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Hospital_Email").Set DataTable("Injury_Hospital_Email","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Hospital_Zip").Set DataTable("Injury_Hospital_Zip","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Hospital_Fax").Set DataTable("Injury_Hospital_Fax","CommonTestData")
		obj_DiactionFrame.WebEdit("Inj_Hospital_Phone").Set DataTable("Injury_Hospital_Phone","CommonTestData")
	End If
	 
	obj_DiactionFrame.WebButton("Next>>").Click
	wait(2)
	
	If ((DataTable("ACC_Fatality","CommonTestData") = "Yes") and (DataTable("CO_Claimant_Dependent","CommonTestData") > 0)) Then
		Dependent()
	End If

End Function
 
Function Dependent()
	
	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Dependent Screen "

	Set obj_InboxPage = Browser("ClaimsBrowser").Page("Inbox")
	Set obj_DiactionFrame = obj_InboxPage.Frame("DIACTION")
	obj_DiactionFrame.WebEdit("Dependent_FirstName").Set DataTable("Dependent_FirstName","CommonTestData")
	obj_DiactionFrame.WebEdit("Dependent_LastName").Set DataTable("Dependent_LastName","CommonTestData")
	obj_DiactionFrame.WebEdit("Dependent_DOB").Set DataTable("Dependent_DOB","CommonTestData")
	obj_DiactionFrame.WebEdit("Dependent_SSN").Set DataTable("Dependent_SSN","CommonTestData")
	obj_DiactionFrame.WebEdit("Dependent_Addr1").Set DataTable("Dependent_Addr1","CommonTestData")
	obj_DiactionFrame.WebEdit("Dependent_Addr2").Set DataTable("Dependent_Addr2","CommonTestData")
	obj_DiactionFrame.WebEdit("Dependent_ZIP").Set DataTable("Dependent_ZIP","CommonTestData")
	obj_DiactionFrame.WebEdit("Dependent_Phone1").Set DataTable("Dependent_Phone1","CommonTestData")
	obj_DiactionFrame.WebList("Dependent_RelationCode").Select DataTable("Dependent_RelationCode","CommonTestData")
	obj_DiactionFrame.WebButton("Next>>").Click
	

End Function

Function Witness()

	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Witness Screen "
	Set obj_InboxPage = Browser("ClaimsBrowser").Page("Inbox")
	Set obj_DiactionFrame = obj_InboxPage.Frame("DIACTION")
	obj_DiactionFrame.WebList("WitnessList").Select DataTable("Witness_List","CommonTestData")
	If DataTable("Witness_List","CommonTestData") = "Yes" Then	
		obj_DiactionFrame.WebEdit("Wit_FirstName").Set DataTable("Witness_FirstName","CommonTestData")
		obj_DiactionFrame.WebEdit("Wit_LastName").Set DataTable("Witness_LastName","CommonTestData")	
		obj_DiactionFrame.WebEdit("Witness_Address1").Set DataTable("Witness_Address1","CommonTestData")	
		obj_DiactionFrame.WebEdit("Witness_Address2").Set DataTable("Witness_Address2","CommonTestData")	
		obj_DiactionFrame.WebEdit("Wit_City").Set DataTable("Witness_City","CommonTestData")					
		obj_DiactionFrame.WebEdit("Wit_Zip").Set DataTable("Witness_Zip","CommonTestData")
		obj_DiactionFrame.WebEdit("html id:=HomePhone.*","html tag:=INPUT").Set DataTable("Witness_PrimaryPhone","CommonTestData")
		obj_DiactionFrame.WebEdit("html id:=Fax.*","html tag:=INPUT").Set DataTable("Witness_Fax","CommonTestData")
	End If
	 
	obj_DiactionFrame.WebButton("Next>>").Click
	obj_InboxPage.Sync
	

End Function

Function Attorney()

	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Attorney Screen "

	Set obj_InboxPage = Browser("ClaimsBrowser").Page("Inbox")
	Set obj_DiactionFrame = obj_InboxPage.Frame("DIACTION")
	obj_DiactionFrame.WebList("AttorneyList").Select DataTable("Attorney_List","CommonTestData")
	Wait(2)
	If DataTable("Attorney_List","CommonTestData") = "Yes" Then
		obj_DiactionFrame.WebEdit("Att_FirmName").Set DataTable("Attorney_FirmName","CommonTestData")	
		Wait(1)
		obj_DiactionFrame.WebEdit("Att_FirstName").MiddleClick
		obj_DiactionFrame.WebEdit("Att_FirstName").Set DataTable("Attorney_FirstName","CommonTestData")	
		obj_DiactionFrame.WebEdit("Att_LastName").Set DataTable("Attorney_LastName","CommonTestData")	
		obj_DiactionFrame.WebEdit("Att_Address1").Set DataTable("Attorney_Address1","CommonTestData")
		obj_DiactionFrame.WebEdit("Att_Address2").Set DataTable("Attorney_Address2","CommonTestData")	
		obj_DiactionFrame.WebEdit("Att_City").Set DataTable("Attorney_City","CommonTestData")	
		obj_DiactionFrame.WebEdit("Att_ZIP").Set DataTable("Attorney_ZIP","CommonTestData")	
		obj_DiactionFrame.WebEdit("Att_Email").Set DataTable("Attorney_Email","CommonTestData")	
	End If
	obj_DiactionFrame.WebButton("Next>>").Click
	

End Function


Function Additional_Information()
	
	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Additional Information Screen "

	Set obj_DiactionFrame = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	If DataTable("CDF","CommonTestData") = "Yes"	 Then
		obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pCDFValues.*l1.*pFieldValue").Set DataTable("CDF_Field1","CommonTestData")
		obj_DiactionFrame.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pCDFValues.*l2.*pFieldValue").Set DataTable("CDF_Field2","CommonTestData")
	End If
	obj_DiactionFrame.WebEdit("Additional_Note").Set DataTable("AdditionalInfo_Note","CommonTestData")
	obj_DiactionFrame.WebButton("Next>>").Click 
	Browser("ClaimsBrowser").Page("Inbox").Sync
	Wait(20)
		
End Function


Function Assignment()

	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Assignment Screen "
	Set obj_DiactionFrame = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	If DataTable("LetRest","CommonTestData") = "TRUE" Then
		if obj_DiactionFrame.WebCheckBox("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pIncidentDetails.*pReportOnly","type:=checkbox").Exist(5) then
			obj_DiactionFrame.WebCheckBox("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pIncidentDetails.*pReportOnly","type:=checkbox").Set "ON"
		End If
	End If
	obj_DiactionFrame.WebButton("Ass_Save").Click
	If (DataTable("AccidentCode_Override_TPA","CommonTestData") = "TRUE") or (DataTable("Monopolistic_Override","CommonTestData") = "TRUE") Then
		Call ReassignOffice()
	End If
	Wait(6)
	obj_DiactionFrame.WebButton("Get_Claim_Number").Click
	Wait(5)
	Browser("ClaimsBrowser").Page("Inbox").Sync
	If  obj_DiactionFrame.WebButton("No Duplicates Found").Exist(3)Then 
		obj_DiactionFrame.WebButton("No Duplicates Found").Click  
	End If
	Wait(10)
	Call GetClaimNumber()	
	
End Function

Function GetClaimNumber()
	
	Claim_Number= Browser("ClaimsBrowser").Page("Inbox").Frame("Review_Distribution_Frame").WebTable("ClaimNumber_Table").GetROProperty("innertext")
	Claim_Number=Trim(Claim_Number)
	Claim_Number=right(Claim_Number,10)
	Environment.Value("NewClaimNumber") =  Claim_Number & "   " & Environment.Value("SCaseId")
	Print "+++++++++++++++++++++++++++++++++++++ Claim Number is +++++++++++++++ " & Environment.Value("NewClaimNumber")  & " ++++++++++++++++++++++++++++++++++++++++++++++++++++++++ "
	
End Function


Function Review_Distribution()
	
	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Review Distribution Screen "
	
	On Error Resume Next

	If (DataTable("AccidentCode_Override_TPA","CommonTestData") = "FALSE") or (TPA_override=True) Then
		Call GetClaimNumber()
		Browser("name:=CCC.*").Page("title:=CCC.*").Sync
		If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			''Log Off	
		Else
			If (DataTable("AccidentCode_Override_TPA","CommonTestData") = "FALSE") or (TPA_override=True) Then
				Set Obj_Page = Browser("name:=CCC.*").Page("title:=CCC.*")
				Set obj_ActionIFrame = Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame")
				Obj_Page.Sync
				obj_ActionIFrame.WebButton("html id:=RLAdd","html tag:=BUTTON").Click
				wait(1)
				'Descriptive object to identify  Web List objects in Review Distribution Screen.
				Set DList=description.Create
				DList("micclass").value="WebList" 
				Set  Obj_WebList = obj_ActionIFrame.ChildObjects(DList) 
				NoOfWebListObj = Obj_WebList.Count
				For Counter=0 to NoOfWebListObj-1 
			     	If Right(Obj_WebList(Counter).getroproperty("name"),15) = "pdistMethodName" then
			     		ChannelValue = Obj_WebList(Counter).getroproperty("value")
			     		CommmonValue = Left(Obj_WebList(Counter).getroproperty("name"),40)
			     		ActalValue = Obj_WebList(Counter).getroproperty("name")
			     		If ChannelValue = "Email" OR ChannelValue = "ELECACK" Then
			     			Set  DEmail=description.Create
							DEmail("micclass").value="WebEdit"
							DEmail("name").value= CommmonValue & "$pemailDtl"
							DEmail("html id").value= "emailDtl"
							DEmail("name").RegularExpression = false
			     			obj_ActionIFrame.WebEdit(DEmail).Set "test@test.com"
			     		ElseIf ChannelValue = "Fax" Then
			     			Set DFax=description.Create
							DFax("micclass").value="WebEdit"
							DFax("name").value= CommmonValue & "$pFax"
							DFax("name").RegularExpression = false
			     			obj_ActionIFrame.WebEdit(DFax).Set "343-334-3434"
			     		ElseIf ChannelValue = "Mail" Then
			     			Set  DAddr1=description.Create
							DAddr1("micclass").value="WebEdit"
							DAddr1("name").value= CommmonValue & "$paddr1Dtl"
							DAddr1("html id").value= "addr1Dtl"
							DAddr1("name").RegularExpression = false
			     			obj_ActionIFrame.WebEdit(DAddr1).Set "Address1"
			     			Set  DZip=description.Create
							DZip("micclass").value="WebEdit"
							DZip("name").value= CommmonValue & "$pzipDtl"
							DZip("html id").value= "zipDtl"
							DZip("name").RegularExpression = false
							obj_ActionIFrame.WebEdit(DZip).Set ""
			     			obj_ActionIFrame.WebEdit(DZip).Set "60196"
			     			Set WshShell = CreateObject("WScript.Shell")
							Wait(1)
							WshShell.SendKeys "{TAB}"
							Wait(2)
							Set WshShell = Nothing
						ElseIf ChannelValue = "#0" Then
							Set  DChannel=description.Create
							DChannel("micclass").value="WebList"
							DChannel("name").value= ActalValue
							DChannel("html id").value= "distMethodName"
							DChannel("name").RegularExpression = false
			     			obj_ActionIFrame.WebList(DChannel).Select "Email"
							
							Set  DCheckbox1=description.Create
							DCheckbox1("micclass").value="WebCheckBox"
							DCheckbox1("name").value= CommmonValue & "$plossNoticeInd"
							DCheckbox1("type").value= "checkbox"
							DCheckbox1("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebList(DCheckbox1).Set "ON"
			     			
			     			Set  DCheckbox2=description.Create
							DCheckbox2("micclass").value="WebCheckBox"
							DCheckbox2("name").value= CommmonValue & "$pcoverLetterInd"
							DCheckbox2("type").value= "checkbox"
							DCheckbox2("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebList(DCheckbox2).Set "ON"
			     			
			     			Set  DCheckbox3=description.Create
							DCheckbox3("micclass").value="WebCheckBox"
							DCheckbox3("name").value= CommmonValue & "$pparListInd"
							DCheckbox3("type").value= "checkbox"
							DCheckbox3("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebList(DCheckbox3).Set "ON"
			     			
			     			Set  DCheckbox4=description.Create
							DCheckbox4("micclass").value="WebCheckBox"
							DCheckbox4("name").value= CommmonValue & "$pcustomFieldsInd"
							DCheckbox4("type").value= "checkbox"
							DCheckbox4("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebList(DCheckbox4).Set "ON"
			     			
			     			Set  DCheckbox5=description.Create
							DCheckbox5("micclass").value="WebCheckBox"
							DCheckbox5("name").value= CommmonValue & "$pnotesInd"
							DCheckbox5("type").value= "checkbox"
							DCheckbox5("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebList(DCheckbox5).Set "ON"
			     			
			     			Set  DCheckbox6=description.Create
							DCheckbox6("micclass").value="WebCheckBox"
							DCheckbox6("name").value= CommmonValue & "$poriginalDocumentInd"
							DCheckbox6("type").value= "checkbox"
							DCheckbox6("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebList(DCheckbox6).Set "ON"
						else
							'Do Nothing
			     	End if
			     End if
		  	Next 
		  	'Descriptive object to identify  Web Edit objects in Review Distribution Screen.
			Set DEdit=description.Create
			DEdit("micclass").value="WebEdit" 
			Set  Obj_WebEdit = obj_ActionIFrame.ChildObjects(DEdit) 
			NoOfWebEditObj = Obj_WebEdit.Count
			For Counter1=0 to NoOfWebEditObj-1 
			     if Right(Obj_WebEdit(Counter1).getroproperty("name"),9) = "pemailDtl" then
			     	If Obj_WebEdit(Counter1).getroproperty("value") = ""Then
			     		Set DEmail1=description.Create
						DEmail1("micclass").value="WebEdit"
						DEmail1("name").value= Obj_WebEdit(Counter1).getroproperty("name")
						DEmail1("html id").value= "emailDtl"
						DEmail1("name").RegularExpression = false
		     			obj_ActionIFrame.WebEdit(DEmail1).Set "test@test.com"
			     	End If
			     End If
			Next	
			If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("xpath:=//button[@title='Complete']").Exist(5) Then
				Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("xpath:=//button[@title='Complete']").Click
			End If
			End If 
		End If 	
		
	Else 
		Call Override_TPA()
		TPA_override = True
	End If

End Function

Function WC_eZAccess()
	
		Environment.value("str_ScreenName") = " EZAccess >>>> Login Screen "
		
		SystemUtil.Run "iexplore.exe", Environment.Value("EZ_URL")
		Wait(3)
		'Open eZAccess QA Link and Maximize the Browser
		'Cleanup the Sessions
		Set Obj_eZAccessStartPage = Browser("title:=eZACCESS Start Page - Internet Explorer").Page("title:=eZACCESS Start Page")
		Obj_eZAccessStartPage.Link("name:=Cleanup Sessions","text:=Cleanup Sessions").Click
		Browser("title:=eZACCESS Start Page - Internet Explorer").Dialog("text:=Message from webpage").WinButton("text:=OK","regexpwndtitle:=OK").Click
		Wait(2)
		'Clicking on Production Sysyem link
		Obj_eZAccessStartPage.Link("name:=eZACCESS Production System","text:=eZACCESS Production System").Click
		'EZAccess Login
		Set Obj_eZAccessBrowserPage = Browser("title:=QA: Zurich Intranet Login - Internet Explorer").Page("title:=QA: Zurich Intranet Login")
		Obj_eZAccessBrowserPage.Sync
		Obj_eZAccessBrowserPage.WebEdit("name:=username","type:=text","html tag:=INPUT").Set Environment.Value("EZ_LoginId")
		Obj_eZAccessBrowserPage.WebEdit("name:=password","type:=password","html tag:=INPUT").Set Environment.Value("EZ_LoginPassword")
		Obj_eZAccessBrowserPage.WebButton("name:=Log In","innertext:=Log In","html tag:=BUTTON","type:=submit").Click
		'Clicking on Claims Search 
		Set Obj_eZAccessACTIIPage =Browser("title:=ACT II - Internet Explorer").Page("title:=ACT II")
		Obj_eZAccessACTIIPage.Frame("name:=topFrame","html tag:=FRAME").WebElement("innertext:=eZACCESS","html id:=ezaMenu_ezaccess").WaitProperty "Visible","True",1000
		Setting.WebPackage("ReplayType") = 2
		Obj_eZAccessACTIIPage.Frame("name:=topFrame","html tag:=FRAME").WebElement("innertext:=eZACCESS","html id:=ezaMenu_ezaccess").FireEvent "onmouseover"
		Setting.WebPackage("ReplayType") = 1
		Obj_eZAccessACTIIPage.Frame("name:=topFrame","html tag:=FRAME").WebElement("innertext:=Claim Search").Click
	    Obj_eZAccessACTIIPage.Frame("name:=sidebarFrame").WebEdit("name:=claimId").Set Claim_Number
		Obj_eZAccessACTIIPage.Frame("name:=sidebarFrame").WebButton("name:=Go!","value:=Go!").Click
		If Browser("ACT II").Page("ACT II").Frame("sidebarFrame").Link("Claim Data").Exist(10) Then
			Environment.Value("ClaimRegistration") = "YES"
			EzRegStatus = True
			Browser("ACT II").Page("ACT II").Frame("sidebarFrame").Link("Claim Data").Click
		Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Claim Summary").Click
		accReportDate = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_LossDate").GetROProperty("value")
		Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Insured").Click
		
		EzAcc_PolicyNum = Browser("Customer_Browser").Page("ACT II").Frame("contentFrame").WebEdit("EZAccesss_PolicyNumber").GetROProperty ("value")
		Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Accident").Click
		lossDate = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("lossDate").GetROProperty("value")		
		accTime = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_Time").GetROProperty("value")
		accDesc = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Acc_Description").GetROProperty("value")
		accCode1 = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("Acc_Code").GetROProperty("value")
		accCode = Split(accCode1, "-")
		agentLoss = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("AgentOfLoss").GetROProperty("value")
		lossLocation = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("dtlLossLocation").GetROProperty("value")
		accState = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("Acc_State").GetROProperty("value")
		Browser("ACT II").Page("ACT II").Frame("sidebarFrame").Link("Claimants").Click
		Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Claimant Detail").Click
		clmtLastName = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("clmt_LastName").GetROProperty("value")
		Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Injury").Click
		injDesc = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Inj_Description").GetROProperty("value")
		injNature = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("Inj_Nature").GetROProperty("value")
		injBodyPart = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("Inj_Bodypart").GetROProperty("value")
		regOccu = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("InjEmp_RegularOccup").GetROProperty("value")
		injOccu = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("InjEmp_InjuryOccup").GetROProperty("value")
		Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Employer").Click
		empLastName = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Emp_Lastname").GetROProperty("value")
		
		If DataTable("ClaimType","CommonTestData") <> "Accident and Health" Then
			Browser("ACT II").Page("ACT II").Frame("contentFrame").Link("Employment").Click
			empHours = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebEdit("Emp_Hours/Day").GetROProperty("value")
			empDays = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("Emp_Days/Week").GetROProperty("value")
			empLostTime = Browser("ACT II").Page("ACT II").Frame("contentFrame").WebList("Emp_LostTime").GetROProperty("value")
		End If

		accDate = Split(DataTable("CS_Accident_Date","CommonTestData"), "/")
		repDate = Split(Trim(accReportDate), "-")
		
	End If	

		
End Function


Function ReassignOffice()

	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Reassign Office Screen "
	wait(3)
	Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*").WebButton("name:=Reassign Office").Click
	Browser("name:=Srchssignment").Page("title:=Srchssignment").WebEdit("name:=.*PTempAssignmentPage.*pTargetCode").Set "10NLT"
	Browser("name:=Srchssignment").Page("title:=Srchssignment").WebButton("name:=Search").Click
	Set obj = Browser("name:=Srchssignment").Page("micClass:=Page")
	Set objWebElement =  obj.webtable("column names:=Assignment;Kind;Name;Name1;Code").ChildItem(2,0,"webelement",0)
	Setting.WebPackage("ReplayType") = 2
	objWebElement.FireEvent "ondblclick",,,micLeftBtn 
	Setting.WebPackage("ReplayType") = 1 
	Browser("name:=Srchssignment").Page("title:=Srchssignment").WebButton("name:=Select").Click

End Function

Function Extract_SCaseId ()

	SCase_Id=""
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("S_Case_ID").Exist(2) Then
		SCase_Id = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("S_Case_ID").GetROProperty ("innertext")
		Print "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  " & SCase_Id & "  +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
	End If
	Environment.Value("SCaseId") = SCase_Id 

End Function



Function StateQuestion_California()

	
	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> California State Question Screen "
	Set obj_DiactionFrame = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	obj_DiactionFrame.WebEdit("State_Ques_CAL_CompLocation").Set DataTable("Edit1","StateQuestions")
	obj_DiactionFrame.WebEdit("State_Ques_CAL_UnempInsurNumber").Set DataTable("Edit2","StateQuestions")
	obj_DiactionFrame.WebList("State_Ques_CAL_CompanyType").Select DataTable("Select1","StateQuestions")
	obj_DiactionFrame.WebList("State_Ques_CAL_EmpTime1").Select DataTable("Select2","StateQuestions")
    obj_DiactionFrame.WebList("State_Ques_CAL_EmpTime2").Select DataTable("Select3","StateQuestions")
	obj_DiactionFrame.WebRadioGroup("State_Ques_CAL_EmpTimeAMPM").Select DataTable("Radio1","StateQuestions")
	obj_DiactionFrame.WebList("State_Ques_CAL_EmpPaidFullWages").Select DataTable("Select4","StateQuestions")
	obj_DiactionFrame.WebEdit("State_Ques_CAL_DateOfEmpClaimForm").Set DataTable("Edit3","StateQuestions")
	obj_DiactionFrame.WebList("StateQues_CO_SalContinued").Select DataTable("Select6","StateQuestions")
	obj_DiactionFrame.WebList("State_Ques_CAL_OtherEmpInjured").Select DataTable("Select5","StateQuestions")
	obj_DiactionFrame.WebEdit("State_Ques_CAL_ClassCodePolicy").Set  DataTable("Edit4","StateQuestions")
 	obj_DiactionFrame.WebButton("Next>>").Click

End Function

Function StateQuestion_North_Carolina()

	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> North Carolina State Question Screen "
	Set obj_DiactionFrame = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	obj_DiactionFrame.WebEdit("StateQues_NC_Full_Wages").Set DataTable("Edit1","StateQuestions")
	obj_DiactionFrame.WebEdit("StateQues_NC_Advantage_Amount").Set DataTable("Edit2","StateQuestions")
	obj_DiactionFrame.WebList("StateQues_NC_Hour").Select DataTable("Select1","StateQuestions")
	obj_DiactionFrame.WebList("StateQues_NC_Min").Select DataTable("Select2","StateQuestions")
	obj_DiactionFrame.WebRadioGroup("State_Ques_NS_TimeAMPM").Select DataTable("Radio1","StateQuestions")
	obj_DiactionFrame.WebEdit("StateQues_NC_Emp_Salary").Set DataTable("Edit5","StateQuestions")
	obj_DiactionFrame.WebList("StateQues_NC_Advantage").Select DataTable("Select3","StateQuestions")
	obj_DiactionFrame.WebEdit("StateQues_NC_Title").Set DataTable("StateQues_NC_Title","StateQuestions")
	obj_DiactionFrame.WebEdit("StateQues_NC_Other_Advantage").Set DataTable("StateQues_NC_Other_Advantage","StateQuestions")
	obj_DiactionFrame.WebButton("Next>>").Click
	
	
End Function

Function StateQuestion_Colorado()

	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Colorado State Question Screen "
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_CO_Self-Insured").Select DataTable("Select1","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_CO_RecevieTips").Select DataTable("Select2","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_CO_FullWagePaid").Select DataTable("Select4","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_CO_ReceiveMeals").Select DataTable("Select5","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pPegaIncidentId.*pClaimants.*pClaimantDet.*l5.*pAnswer").Select DataTable("Select5","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_CO_ReceiveRoom").Select DataTable("Select7","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pPegaIncidentId.*pClaimants.*pClaimantDet.*l7.*pAnswer").Select DataTable("Select5","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_CO_HealthInsIncluded").Select DataTable("Select9","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_CO_TempDisbled").Select DataTable("Select10","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_CO_Time1").Select DataTable("Select11","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_CO_Time2").Select DataTable("Select12","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("StateQues_CO_AM_PM").Select DataTable("Radio1","StateQuestions")
	Browser("title:=CCC Manager.*").Page("title:=CCC Manager.*").Frame("name:=Pega.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pPegaIncidentId.*pClaimants.*pClaimantDet.*l11.*pAnswer").set DataTable("Edit1","StateQuestions")'WebEdit("StateQues_CO_EmpDoing").Set DataTable("Edit1","StateQuestions")
	Browser("title:=CCC Manager.*").Page("title:=CCC Manager.*").Frame("name:=Pega.*").WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pPegaIncidentId.*pClaimants.*pClaimantDet.*l13.*pAnswer").Select DataTable("Select13","StateQuestions")
	Browser("title:=CCC Manager.*").Page("title:=CCC Manager.*").Frame("name:=Pega.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pPegaIncidentId.*pClaimants.*pClaimantDet.*l14.*pAnswer").Set DataTable("Edit2","StateQuestions")
	Browser("title:=CCC Manager.*").Page("title:=CCC Manager.*").Frame("name:=Pega.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pPegaIncidentId.*pClaimants.*pClaimantDet.*l15.*pAnswer").Set DataTable("Edit3","StateQuestions")
	Browser("title:=CCC Manager.*").Page("title:=CCC Manager.*").Frame("name:=Pega.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pPegaIncidentId.*pClaimants.*pClaimantDet.*l12.*pAnswer").Set DataTable("Edit4","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	

End function

Function StateQuestion_NewYork()

	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> NewYork State Question Screen "

	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1_2").Frame("PegaGadget0Ifr").WebEdit("test1").Set DataTable("Edit1","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_NY_Full_WagePaid").Select DataTable("Select8","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_NY_BWnum").Set DataTable("Edit1","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pPegaIncidentId.*pClaimants.*pClaimantDet.*l2.*pAnswer").Object.value = DatTable("Edit2","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	
End function

Function StateQuestion_Oklahoma()
	
	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Oklahoma State Question Screen "
		
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_Oklaha_YearsInjurEmployee").Set DataTable("Edit1","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_Oklaha_MonthsInjurEmployee").Set DataTable("Edit2","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_Oklaha_EmploymentAgreement").Select DataTable("Select1","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_Oklaha_Hour").Select DataTable("Select2","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_Oklaha_Minute").Select DataTable("Select3","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("StateQues_Oklaha_Begin_Work_AM").Select DataTable("Radio1","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_Oklaha_Injury_Result_from").Select DataTable("Select4","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_Oklaha_Participate_CWMP").Select DataTable("Select5","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_Oklaha_CWMP").Object.value = DataTable("Edit3","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_Oklaha_BType").Object.value =  DataTable("Edit4","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_Oklaha_Ownership").Select DataTable("Select6","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_Oklaha_ReporterName").Object.value =  DataTable("Edit5","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_Oklaha_ReporterTitle").Object.value =  DataTable("Edit6","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_Oklaha_ReporterPhone").Object.value = DataTable("Edit7","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_Oklaha_Date").Set "7/14/2015"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	
	
End Function

Function StateQuestion_Oregon()

	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Oregon State Question Screen "
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_OR_Work_Begin_Hour").Select DataTable("Select1","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_OR_Work_Begin_Min").Select DataTable("Select2","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("StateQues_OR_Begin_AM").Select DataTable("Radio1","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_OR_Work_End_Hour").Select DataTable("Select3","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_OR_Work_End_Min").Select DataTable("Select4","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("StateQues_OR_End_PM").Select DataTable("Radio2","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("StateQues_OR_Mon").Set DataTable("Checkbox1","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_OR_LegalBusinessName").Set DataTable("Edit1","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_OR_Federral_ID").Set DataTable("Edit2","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_OR_Leased_Employee").Set DataTable("Edit3","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_OR_Phone_first_three").Set DataTable("Edit4","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_OR_Phone_middle_three").Set DataTable("Edit5","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_OR_Phone_last_three").Set DataTable("Edit6","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_OR_ClientFederalID").Set DataTable("Edit7","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("StateQues_OR_ClassCode").Set DataTable("Edit8","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_OR_Accident_Cause").Select DataTable("Select5","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("StateQues_OR_OtherInjured").Select DataTable("Select6","StateQuestions")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click

End function
	

Function Logout()
	
	Environment.value("str_ScreenName") = "Carepoint - WC  >>>> Logoff Screen "
	
	Browser("name:=CC.*").Page("title:=CC.*").Image("name:=Image","image type:=Image Link","Index:=0").Click
	Browser("name:=CC.*").Page("title:=CC.*").WebElement("innertext:=Log off","html id:=ItemMiddle").Click
	SystemUtil.CloseProcessByName "iexplore.exe"
	
End Function
