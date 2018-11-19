'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
								'Carepoint CD Business Functions
								'Updated By : Srirekha Talasila
								'Updated On : 12/19/2016
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Overide_TPA = 0

 Function Login()

	Environment.value("str_ScreenName") = "Carepoint >>>> CD - Login Page "
	SystemUtil.CloseProcessByName "iexplore.exe"
	Systemutil.Run "iexplore.exe", Environment.Value("CP_URL")	
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("UserIdentifier").Set Environment.Value("CP_LoginId")
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("Password").Set Environment.Value("CP_LoginPassword") 
	Browser("ClaimsBrowser").Page("LoginPage").WebButton("Log In").Click
	 
 End function

Function Select_CD()
	
	Environment.value("str_ScreenName") = "Carepoint >>>> CD - Select CD "
	Browser("name:=CCC.*").Page("title:=CCC.*").Link("html tag:=A","name:=My Group").Click
	Browser("name:=CCC.*").Page("title:=CCC.*").WebList("html id:=objWorkBasketSelect","html tag:=SELECT","name:=select").Select "Construction Defect"
		
End function


Function Select_Workitem()

	Environment.value("str_ScreenName") = "Carepoint >>>> CD - Select WorkItem "
'	Browser("Customer_Browser").Page("WorkList_Basket").WebElement("SortDate").Click
	wait(3)
	Browser("Customer_Browser").highlight
	Browser("Customer_Browser").Page("WorkList_Basket").WebElement("title:=Click.*","Index:=12").click
'	Browser("Customer_Browser").Sync
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
							Wait(3)
						End If
					End If
				Next
				Set tabobj=nothing
				SelectionCount=SelectionCount+1	
		Else
			rem SelectionCount=SelectionCount+1
         End If
		
			Check=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Customer Search").Exist(3)
			 If Check="True" Then
					Exit Do
			End If
	Loop Until Check=False
	
	' Clicking the Customer Search Button
	 
	If Browser("title:=TestDaeja.*").Exist(3) Then
		Browser("title:=TestDaeja.*").Close 
	    Wait(1)
	End If
	
	Customer_Search()
	
End Function


Function Customer_Search()

	Environment.value("str_ScreenName") = "Carepoint >>>> CD - Customer Search "
	Dim objBrwpage_CustomerSearch

	set objBrwpage_CustomerSearch=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame") 
	objBrwpage_CustomerSearch.WebButton("Customer Search").Click 
	Wait(3)
	If (DataTable("Search_Flag","GL-Data") = "Customer") Then
		If objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Exist(10) Then
		   objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Set CDATE(DataTable("CS_Accident_Date","GL-Data"))	
		End If		
		objBrwpage_CustomerSearch.WebEdit("CS_CustomerName").Set DataTable("CS_Customer_Name","GL-Data")
		objBrwpage_CustomerSearch.WebEdit("CS_SiteCode").Set DataTable("CS_SiteCode","GL-Data")
		objBrwpage_CustomerSearch.WebButton("CS_Search").Click	
		Browser("ClaimsBrowser").Sync	
		Index=1
		while index<>0'''Here the condition will waits till Web Table load
			If (objBrwpage_CustomerSearch.webelement("CS_No_Matching_Data").Exist(10) ) Then
				index=0
			Else
				index=0
				Set obj_BusinessUnit=Browser("CreationTime:=0").Page("title:=.*").Frame("name:=actionIFrame").WebTable("column names:=Click to sortBusiness Unit ,;Click to sortCustomer Name ,;Click to sortEntity Name ,;Click to sortSite Name ,;Click to sortSite Code ,;Click to sortAddress 1 ,;Click to sortAddress 2 ,;Click to sortCity ,;Click to sortState ,;Click to sortZip Code ,;Click to sortPhone ,;Click to sortFax ,","index:=23").ChildItem(2,1,"WebElement",0)''@DP
				If obj_BusinessUnit.Exist(30) Then
					obj_BusinessUnit.click '''This will target first row in the Customer SEarch result 
				End If
				wait(1)
				If objBrwpage_CustomerSearch.WebButton("html id:=submitButton").Exist(20) Then
					Setting.WebPackage("ReplayType") = 1
					objBrwpage_CustomerSearch.WebButton("html id:=submitButton").Click
					Setting.WebPackage("ReplayType") = 2
				End If
				
				If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(5) then
					Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click				
				End if
			End If 	
		Wend
		
		
		If Browser("title:=Care.*").Exist(3) Then
		   Browser("title:=Care.*").Close 
   		End If  	
   		
		If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Click
			Browser("ClaimsBrowser").Sync
		End If
		If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(5) then
			Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click
		Else
			If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist(5) then
				Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
			End If 
		End If
	ElseIf DataTable("Search_Flag","GL-Data") = "Employee" Then		
		Employee_Search()
	Else
		Add_NewCustomer()
	End If
	
End Function


Function Add_NewCustomer()

		Environment.value("str_ScreenName") = "Carepoint - CD >>>> Add New Customer "
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Add_New_Customer").Click
		Browser("ClaimsBrowser").sync
        Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Name").Set DataTable("AddCust_Name","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Address").Set DataTable("AddCust_Address","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Zip").Set DataTable("AddCust_Zip","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Phone").Set DataTable("AddCust_Phone","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Email").Set DataTable("AddCust_Email","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_SiteCode").Set DataTable("CS_SiteCode","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("AddCust_Submit").Click
		Browser("ClaimsBrowser").Sync
		If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(3) then
				Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click
		Else
			If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist(3) then
				Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
			End If 
		End If
		
		If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Click
		End If
		
End Function


Function Employee_Search()

	Environment.value("str_ScreenName") = "Carepoint - CD >>>>  Employee Search "
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Employee Search").Click
	Browser("ClaimsBrowser").Sync
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Emp_CustomerName").Exist(8) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Emp_CustomerName").Select DataTable("Emp_CustomerName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Emp_Search").Click
	End If
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebRadioGroup("Emp_Result").Exist(15) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebRadioGroup("Emp_Result").Select "1"
	End If
    Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Emp_Select").Click
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Click
	End If
	If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist(5) then
		Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
	End If 

End Function


Function Extract_SCaseId ()

	SCase_Id=""
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("SCase_Id").Exist(2) Then
		SCase_Id = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("SCase_Id").GetROProperty ("innertext")
		Print "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  " & SCase_Id & "  +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
	End If
	Environment.Value("SCaseId") = SCase_Id 

End Function


Function Incident() 
	
	Dim objBrwpage_AddCusmtomer 
	
	Environment.value("str_ScreenName") = "Carepoint - CD >>>>  Incident Screen "
	Browser("ClaimsBrowser").Sync
	set objBrwpage_Incident=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	objBrwpage_Incident.WebEdit("IN_AccidentDate").Set DataTable("CS_Accident_Date","GL-Data")
	objBrwpage_Incident.WebList("AccidentState").Select DataTable("IN_AccidentState","GL-Data")
	objBrwpage_Incident.WebList("AccidentTime1").Select DataTable("IN_AccidentTime1","GL-Data")
	objBrwpage_Incident.WebList("AccidentTime2").Select DataTable("IN_AccidentTime2","GL-Data")
	objBrwpage_Incident.WebList("AccidentTime3").Select DataTable("IN_AccidentTime3","GL-Data")
	objBrwpage_Incident.WebEdit("AccidentDescription").Set DataTable("IN_AccDescription","GL-Data")
	objBrwpage_Incident.WebList("Catagory").Select DataTable("IN_Category","GL-Data")
	objBrwpage_Incident.WebEdit("CD_Incident_Lname").Set DataTable ("IN_CD_Lname","GL-Data")
 	objBrwpage_Incident.WebList("MasterTrailer_Field").Select DataTable ("Master_and_Trailer_Field","GL-Data")
	objBrwpage_Incident.WebButton("Next>>").Click
	'If Duplicate Claim Exists
	If objBrwpage_Incident.WebButton("No Duplicates Found").Exist(5) Then
		objBrwpage_Incident.WebButton("No Duplicates Found").Click
	Else 
		'Do Nothing
	End If
	
	
End Function


Function PolicySearch()

	Dim objBrwpage_PolicySearch 
	Environment.value("str_ScreenName") = "Carepoint - CD >>>>  Policy Search "
	Browser("ClaimsBrowser").Sync
	set objBrwpage_PolicySearch=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	cell_data =objBrwpage_PolicySearch.WebTable("Policy_Table").GetCellData(2,1)  
	
	If cell_data = "" AND DataTable("CS_Policynum","GL-Data")="" Then
		Set polobj = objBrwpage_PolicySearch.WebTable("Policy List")
		Set polobj2 = polobj.ChildItem(2,1,"WebRadioGroup",0)			
		d = polobj2.getroproperty("class")
		If d = "Radio" Then 
			objBrwpage_PolicySearch.WebRadioGroup("Policy_RadioButton").Click
			objBrwpage_PolicySearch.WebButton("Next>>").Click
		Else
			'Report to html "No Policy Found
		End if
	End if
	pol_flag = False
	If objBrwpage_PolicySearch.WebElement("Nomatchingpolicy").Exist(3) or DataTable("CS_Policynum","GL-Data") <> "" Then  
		pol_flag = True
		PolNum=Trim(DataTable("CS_Policynum","GL-Data"))
		
		objBrwpage_PolicySearch.WebEdit("PS_Policynum").Set PolNum
		objBrwpage_PolicySearch.WebEdit("html id:=zpsMonthsPrior").Set Trim(DataTable("Months_Prior","GL-Data"))
		objBrwpage_PolicySearch.WebButton("Policy_Retrieve").Click
		Browser("ClaimsBrowser").Sync
		cell_data = objBrwpage_PolicySearch.WebTable("Policy_Table").GetCellData(2,1)
		If Not objBrwpage_PolicySearch.WebElement("Nomatchingpolicy").Exist(5) and cell_data=""Then  
			objBrwpage_PolicySearch.WebRadioGroup("Policy_RadioButton").Click
			objBrwpage_PolicySearch.WebButton("Next>>").Click
		Else 
			objBrwpage_PolicySearch.WebRadioGroup("Policy_RadioButton").Select "Indeterminate"
			objBrwpage_PolicySearch.WebButton("Next>>").Click
		End if
	End If
	
End Function


Function Override_TPA()

    If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Exist(5) then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Click
	Else
		'Do Nothing
	End If
	

End Function

Function ReassignOffice()

	Environment.value("str_ScreenName") = "Carepoint - CD  >>>> Reassign Office Screen "
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



Function Contact_Info()
		
	Dim objBrwpage_Contact_Info
	Environment.value("str_ScreenName") = "Carepoint - CD >>>> Contact Info "
	Browser("ClaimsBrowser").Sync
	Set objBrwpage_Contact_Info = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	objBrwpage_Contact_Info.WebEdit("CO_Rep_Name").Set DataTable("CO_Rep_Name","GL-Data")
	objBrwpage_Contact_Info.WebEdit("CO_Rep_Email").Set DataTable("CO_Rep_Email","GL-Data")
	objBrwpage_Contact_Info.WebEdit("CO_Rep_Phone").Set DataTable("CO_Rep_Phone","GL-Data")
	objBrwpage_Contact_Info.WebList("CO_Report_Relation").Select DataTable("CO_Rep_Relation","GL-Data")
	objBrwpage_Contact_Info.WebEdit("CO_CusCon_Email").Set DataTable("CO_CusCon_Email","GL-Data")
	objBrwpage_Contact_Info.WebEdit("CO_CusCon_Fax").Set DataTable("CO_CusCon_Fax","GL-Data")		
	If objBrwpage_Contact_Info.WebEdit("CO_Party_FName").Exist(3) Then
		objBrwpage_Contact_Info.WebEdit("CO_Party_FName").Set DataTable("CO_Party_FName","GL-Data")
	End If
	AttorneyYN= DataTable("CO_AttorneyYN","GL-Data")
	If (AttorneyYN = "Yes") Then
		objBrwpage_Contact_Info.WebEdit("CO_FirmName").Set DataTable("CO_FirmName","GL-Data")
		objBrwpage_Contact_Info.WebEdit("CO_AttPhone").Set DataTable("CO_AttPhone","GL-Data")
		objBrwpage_Contact_Info.WebEdit("CO_AttZip").Set DataTable("CO_AttZip","GL-Data")
		objBrwpage_Contact_Info.WebButton("Next>>").Click
	Else
		objBrwpage_Contact_Info.WebButton("Next>>").Click
		
	End If
	
		
End function


Function Accident_Page()
	
	Dim objBrwpage_Accident_Page
	Environment.value("str_ScreenName") = "Carepoint - CD >>>> Accident Screen "
	Set objBrwpage_Accident_Page = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	Browser("ClaimsBrowser").Sync	
	objBrwpage_Accident_Page.WebList("ACC_AccCode").Select DataTable("ACC_AccCode","GL-Data")
	objBrwpage_Accident_Page.WebList("ACC_AgentLoss").Select DataTable("ACC_AgentLoss","GL-Data")
	objBrwpage_Accident_Page.WebButton("Save").Click 
	Browser("ClaimsBrowser").Sync
	objBrwpage_Accident_Page.WebList("ACC_LossLoc").Select DataTable("ACC_LossLoc","GL-Data")
	objBrwpage_Accident_Page.WebList("ACC_SiteAddress").Select DataTable("ACC_SiteAddress","GL-Data")
	If  DataTable("ACC_SiteAddress","GL-Data") = "No" Then
		objBrwpage_Accident_Page.WebEdit("ACC_AccAddress1").Set DataTable("ACC_AccAddress1","GL-Data")
		objBrwpage_Accident_Page.WebEdit("ACC_AccAddress2").Set DataTable("ACC_AccAddress2","GL-Data")
		objBrwpage_Accident_Page.WebEdit("ACC_AccZip").Set DataTable("ACC_AccZip","GL-Data")
		objBrwpage_Accident_Page.WebEdit("ACC_Comments").Click
	End If
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("title:=FNOL.*").WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAccident.*pAddr.*pState").GetROProperty("value")="Select..." Then ''''no value exist in the Zip code
		Browser("ClaimsBrowser").Page("Inbox").Frame("title:=FNOL.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAccident.*pAddr.*ppostalCode").Set DataTable("ACC_AccZip","GL-Data")
	End If
	objBrwpage_Accident_Page.WebEdit("ACC_Comments").Set DataTable("ACC_Comments","GL-Data")
 	objBrwpage_Accident_Page.WebButton("Next>>").Click 

End function


Function Attorney()
	
	Dim objBrwpage_Attorney
	Environment.value("str_ScreenName") = "Carepoint - CD >>>> Attorney Screen "
	Browser("ClaimsBrowser").Sync
	Set objBrwpage_Attorney = Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Attorney")
	objBrwpage_Attorney.WebList("AttorneyList").Select DataTable("Attorney_List","GL-Data")
	If DataTable("Attorney_List","Property") = "Yes" Then
		objBrwpage_Attorney.WebEdit("Attorney_FirmName").Set DataTable("Attorney_FirmName","GL-Data")
		objBrwpage_Attorney.WebEdit("Attorney_FirstName").Set DataTable("Attorney_FirstName","GL-Data")
		objBrwpage_Attorney.WebEdit("Attorney_LastName").Set DataTable("Attorney_LastName","GL-Data")
		objBrwpage_Attorney.WebEdit("Attorney_Address1").Set DataTable("Attorney_Address1","GL-Data")
		objBrwpage_Attorney.WebEdit("Attorney_ZIP").Set DataTable("Attorney_ZIP","GL-Data")
		objBrwpage_Attorney.WebEdit("Attorney_Email").Set DataTable("Attorney_Email","GL-Data")
		objBrwpage_Attorney.WebEdit("Attorney_Fax").Set DataTable("Attorney_Fax","GL-Data")
	End If
	 
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	
	
End Function

Function Witness()
	
	Dim objBrwpage_Witness
	Environment.value("str_ScreenName") = "Carepoint - CD >>>> Witness Screen "
	Browser("ClaimsBrowser").Sync
	Set objBrwpage_Witness = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	objBrwpage_Witness.WebList("WitnessList").Select DataTable("Witness_List","GL-Data")
	If DataTable("Witness_List","GL-Data") = "Yes" Then	
		objBrwpage_Witness.WebEdit("Wit_FirstName").Set DataTable("Witness_FirstName","GL-Data")
		objBrwpage_Witness.WebEdit("Wit_LastName").Set DataTable("Witness_LastName","GL-Data")
		objBrwpage_Witness.WebEdit("Wit_Address1").Set DataTable("Witness_Address1","GL-Data")
		objBrwpage_Witness.WebEdit("Wit_Zip").Set DataTable("Witness_Zip","GL-Data")
		objBrwpage_Witness.WebEdit("html id:=HomePhone","html tag:=INPUT").Set DataTable("Witness_PrimaryPhone","GL-Data")
		objBrwpage_Witness.WebEdit("html id:=Fax","html tag:=INPUT").Set DataTable("Witness_Fax","GL-Data")
	End If
	objBrwpage_Witness.WebButton("Wit_Next>>").Click

End Function


Function Additional_Information()

	Dim objBrwpage_Additional_Information
	Environment.value("str_ScreenName") = "Carepoint - CD >>>> Additional Info Screen "
	Browser("ClaimsBrowser").Sync
	Set objBrwpage_Additional_Information= Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	
	If  Datatable("Zurich_Insured_Attorney","GL-Data")="Yes" Then
		objBrwpage_Additional_Information.WebList("ZInsured_Attorney").Select DataTable("Zurich_Insured_Attorney","GL-Data") 
		objBrwpage_Additional_Information.WebEdit("ZInsured_Firmname").Set DataTable("Zurich_Insured_Firmname","GL-Data")
		objBrwpage_Additional_Information.WebEdit("ZInsured_Fname").Set DataTable("Zurich_Insured_Fname","GL-Data")
		objBrwpage_Additional_Information.WebEdit("ZInsured_Lname").Set DataTable("Zurich_Insured_Lname","GL-Data")  
		objBrwpage_Additional_Information.WebEdit("ZInsured_Zip").Set DataTable("Attorney_Zip","GL-Data")
		
		If  DataTable("Name_Insured_Claim","GL-Data")="No" Then
			objBrwpage_Additional_Information.WebList("Name_Insured_Claim").Select DataTable("Name_Insured_Claim","GL-Data")
			objBrwpage_Additional_Information.WebEdit("ZAttorney_Notes").Set DataTable("Zurich_Attorney_Notes","GL-Data")
			objBrwpage_Additional_Information.WebList("Zurich_Attorney_Category").Select DataTable("Zurich_Attorney_Category","GL-Data")
		End If
		
		If  DataTable("Name_Insured_Claim","GL-Data")="Yes" Then
			objBrwpage_Additional_Information.WebList("Name_Insured_Claim").Select DataTable("Name_Insured_Claim","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Docs").Set DataTable("Name_Insured_Docs","GL-Data")
			objBrwpage_Additional_Information.WebList("Name_Insured_Scan_Fax_Web_Paper").Select DataTable("Name_Insured_Scan_Fax_Web_Paper","GL-Data")
			objBrwpage_Additional_Information.WebList("Name_Insured_Scan_Fax_Web_Paper").Select DataTable("Name_Insured_Scan_Fax_Web_Paper","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Pages_Docs").Set DataTable("Name_Insured_Pages_Docs","GL-Data")
			objBrwpage_Additional_Information.WebList("Name_Insured_Exact").Select DataTable("Name_Insured_Exact","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Setup_Claim").Set DataTable("Name_Insured_Setup_Claim","GL-Data")					
			objBrwpage_Additional_Information.WebList("Name_Insured_Claim_Setup_AI_NI").Select DataTable("Name_Insured_Claim_Setup_AI_NI","GL-Data")
			objBrwpage_Additional_Information.WebList("Name_Insured_Pol_Clm_Recived").Select DataTable("Name_Insured_Pol_Clm_Recived","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Why").Set DataTable("Name_Insured_Why","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Policy_Used").Set DataTable("Name_Insured_Policy_Used","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Other_Info").Set DataTable("Name_Insured_Other_Info","GL-Data")			
			objBrwpage_Additional_Information.WebEdit("ZAttorney_Notes").Set DataTable("Zurich_Attorney_Notes","GL-Data")
			objBrwpage_Additional_Information.WebList("Zurich_Attorney_Category").Select DataTable("Zurich_Attorney_Category","GL-Data")
		End If
	 
   ElseIf  Datatable("Zurich_Insured_Attorney","GL-Data")="No"   Then
   		objBrwpage_Additional_Information.WebList("ZInsured_Attorney").Select DataTable("Zurich_Insured_Attorney","GL-Data") 
		objBrwpage_Additional_Information.WebEdit("Zurich_Attorney_Notes").Set DataTable("Zurich_Attorney_Notes","GL-Data")
		objBrwpage_Additional_Information.WebList("Zurich_Attorney_Category").Select DataTable("Zurich_Attorney_Category","GL-Data")
		
		If  DataTable("Name_Insured_Claim","GL-Data")="No" Then
			objBrwpage_Additional_Information.WebList("Name_Insured_Claim").Select DataTable("Name_Insured_Claim","GL-Data")
		End If
		
		If  DataTable("Name_Insured_Claim","GL-Data")="Yes" Then
			objBrwpage_Additional_Information.WebList("Name_Insured_Claim").Select DataTable("Name_Insured_Claim","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Docs").Set DataTable("Name_Insured_Docs","GL-Data")
			objBrwpage_Additional_Information.WebList("Name_Insured_Scan_Fax_Web_Paper").Select DataTable("Name_Insured_Scan_Fax_Web_Paper","GL-Data")
			objBrwpage_Additional_Information.WebList("Name_Insured_Scan_Fax_Web_Paper").Select DataTable("Name_Insured_Scan_Fax_Web_Paper","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Pages_Docs").Set DataTable("Name_Insured_Pages_Docs","GL-Data")
			objBrwpage_Additional_Information.WebList("Name_Insured_Exact").Select DataTable("Name_Insured_Exact","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Setup_Claim").Set DataTable("Name_Insured_Setup_Claim","GL-Data")					
			objBrwpage_Additional_Information.WebList("Name_Insured_Claim_Setup_AI_NI").Select DataTable("Name_Insured_Claim_Setup_AI_NI","GL-Data")
			objBrwpage_Additional_Information.WebList("Name_Insured_Pol_Clm_Recived").Select DataTable("Name_Insured_Pol_Clm_Recived","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Why").Set DataTable("Name_Insured_Why","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Policy_Used").Set DataTable("Name_Insured_Policy_Used","GL-Data")
			objBrwpage_Additional_Information.WebEdit("Name_Insured_Other_Info").Set DataTable("Name_Insured_Other_Info","GL-Data")			
		End If
	End If
	 
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync	
			
End Function


Function Assignment()
	
	Environment.value("str_ScreenName") = "Carepoint - CD >>>> Assignment Screen "
	
	Dim objBrwpage_Assignment
	Browser("ClaimsBrowser").Sync
	Set objBrwpage_Assignment= Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	
	If DataTable("TPA_Override","GL-Data")="YES" Then
				Call ReassignOffice()
	End If
	
	If objBrwpage_Assignment.WebButton("Get_Claim_Number").Exist(5) Then
		objBrwpage_Assignment.WebButton("Get_Claim_Number").Click
	End If		
	Browser("ClaimsBrowser").Sync	 	
	If  objBrwpage_Assignment.WebButton("class:=buttonTdButton","html tag:=BUTTON","name:=No Duplicates Found").Exist(5)  Then			
		objBrwpage_Assignment.WebButton("class:=buttonTdButton","html tag:=BUTTON","name:=No Duplicates Found").Click	
	End If
	Browser("ClaimsBrowser").Sync
	Call GetClaimNumber()
		
End Function


Function GetClaimNumber()

	Claim_Number= Browser("ClaimsBrowser").Page("Inbox").Frame("Review_Distribution_Frame").WebTable("ClaimNumber_Table").GetROProperty("innertext")
	Claim_Number=Trim(Claim_Number)
	Claim_Number=right(Claim_Number,10)
	Environment.Value("Claim_Number") = Claim_Number
	Environment.Value("NewClaimNumber") =  Claim_Number & "   " & Environment.Value("SCaseId")
	Print "+++++++++++++++++++++++++++++++++++++ Claim Number is +++++++++++++++ " & Environment.Value("NewClaimNumber")  & " ++++++++++++++++++++++++++++++++++++++++++++++++++++++++ "
	
	
End function

'Created By :-  Srirekha Talasila
'This will handle Distributions in Review Screen 

Function Review_Distribution()
	
		Environment.value("str_ScreenName") = "Carepoint - CD  >>>> Review Distribution Screen "
		On Error Resume Next
		Call GetClaimNumber()
		Browser("name:=CCC.*").Page("title:=CCC.*").Sync
		If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			''Log Off	
		Else
		
			If DataTable("TPA_Override","GL-Data")="YES" and Overide_TPA = 0 Then
				Call Override_TPA()
				Overide_TPA = 1
			Else
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
						End If
			     End If
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
	
 End Function

'Created By :-  Srirekha Talasila
'This will Verify Search Functionality using S-Case and Claim Number
Function Binocular_Search()

	Environment.value("str_ScreenName") = "Carepoint - CD  >>>> Binocular Search Screen "
	Browser("ClaimsBrowser").Page("Inbox").Link("Binocular_Search").Click
	Wait(3)
	Browser("SearchIncident").Sync
	Browser("SearchIncident").Page("SearchIncident").WebEdit("Claim_Number").Set DataTable("Claim_Number","GL-Data")
	Browser("SearchIncident").Page("SearchIncident").WebButton("Search_Btn").Click
	Browser("SearchIncident").Sync
	ClaimNumber = Browser("SearchIncident").Page("SearchIncident").WebTable("Search_Results_Table").GetCellData(2,2)
	
	If TRIM(DataTable("Claim_Number","GL-Data")) = TRIM(ClaimNumber) Then
		Set Obj_Result = Browser("SearchIncident").Page("SearchIncident").WebTable("Search_Results_Table").ChildItem(2,0,"WebElement",0)
		Obj_Result.highlight
		Obj_Result.click
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY CLAIMNUMBER","PASS","Claim Number " &ClaimNumber& " Exist in WebTable")	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY CLAIMNUMBER","FAIL","Claim Number " &ClaimNumber& " NOT Exist in WebTable")			
	End If
	Browser("SearchIncident").Page("SearchIncident").WebButton("Clear_Btn").Click
	Browser("SearchIncident").Page("SearchIncident").WebEdit("Incident_Num").Set DataTable("SCase","GL-Data")
	Browser("SearchIncident").Page("SearchIncident").WebRadioGroup("S-Case-Include").Select "Include"
	Browser("SearchIncident").Page("SearchIncident").WebButton("Search_Btn").Click
	Browser("name:=Claim CC Service.*").Sync
	SNumber = Browser("Claim CC Service Items").Page("Claim CC Service Items").WebTable("Claim_Number").GetCellData(2,1)
	If TRIM(DataTable("SCase","GL-Data")) = TRIM(SNumber) Then
		Set Obj_Result1 = Browser("Claim CC Service Items").Page("Claim CC Service Items").WebTable("Claim_Number").ChildItem(2,0,"WebElement",0)
		Obj_Result1.highlight
		Obj_Result1.click
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY S-CASE","PASS","S-CASE " &SNumber& " Exist in WebTable")	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY S-CASE","FAIL"," S-CASE " &SNumber& " NOT Exist in WebTable")			
	End If
	Browser("name:=Claim CC Service.*").Close

End Function 



Function Void_Incident()

	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Void Incident").Click
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Enter_VoidReason").Select DataTable("Enter_VoidReason","GL-Data")
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("VI_Reason").Exist Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("VI_Reason").Set DataTable("VI_Reason","GL-Data")
	End If
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("VI_Submit").Click
     Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Reason_Code").Select DataTable("Reason_Code","GL-Data")
     Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("ReasonCode_Submit").Click
	 Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click

End Function


Function Property_Damage()
		
		Environment.value("str_ScreenName") = "Carepoint - CD  >>>> Property Damage Screen "
		Browser("ClaimsBrowser").Sync
    	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PropertyDam1_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
    	If DataTable("PropertyDam_Location","GL-Data") = "O" Then
    		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=addressLines","Index:=2").Set "Address1"
    		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=addressLines","Index:=3").Set "Address2"
    		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=1").Set "12345"
    	End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropertyDam1_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")

		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist Then
                Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Click
		Else 
			'Do Nothing
		End If

End Function

Function Logout()
	
	Environment.value("str_ScreenName") = "Carepoint - CD  >>>> Logoff Screen "
	
	Browser("name:=CC.*").Page("title:=CC.*").Image("name:=Image","image type:=Image Link","Index:=0").Click
	Browser("name:=CC.*").Page("title:=CC.*").WebElement("innertext:=Log off","html id:=ItemMiddle").Click
	SystemUtil.CloseProcessByName "iexplore.exe"
	
End Function


Function TC16_UC114_Close_and_Reselect_Customer_Property_Damage1()

	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Close").Click
	wait 5

	If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist then
		Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("Cancel").Click
	Else
	End if


	If Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebList("select_work_type").Exist Then
		ReportResult_Event micPass, "Invoking Business component: TC16_UC114_Close_and_Reselect_Customer_Property_Damage1" , "Page is navigated to Inbox after Cancel button is clicked from the Popup"
		Status = trim(Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebTable("workitems").GetCellData(2,2))
		ScaseID = trim(Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebTable("workitems").GetCellData(2,7))
		
			If Status = "Pending" and ScaseID = Environment.Value("SCaseId") Then 
			
			Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebElement("Incident_ID").Click
		
				ReportResult_Event micPass, "Invoking Business component: TC16_UC114_Close_and_Reselect_Customer_Property_Damage1" , "WorkItem with the Status = 'Pending' and IncidentID ="& Environment.Value("SCaseId") & "is present in the Inbox page"
					If  Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebButton("Re-select Customer").Exist Then
						ReportResult_Event micPass, "Invoking Business component: TC16_UC114_Close_and_Reselect_Customer_Property_Damage1" , "Page is navigated to PropertyDamage after selecting the pending workitem from the Inbox page"
							Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebButton("Re-select Customer").Click
								If Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebElement("Customer Search").Exist Then
								Else
					End If
					Else
					End If
			Else
			End If

	Else
	End If

End Function




Function IncidentLink()

	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").Link("Incident_Link").Click
	Wait(3)

End Function





Function TC08_E2E_ReselectCustomer_scenario()
	
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("AddInfo_ReSelect_Cust").Click
	Wait 2
	Dim counter
	counter = Environment.Value("counter")
	counter = counter + 1
	DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
	Environment.Value("TC08_E2E_ReselectCustomer_scenario")=True
	Call Customer_Search()
 	

End Function


Function TC09_E2E_ReselectEmployee_scenario()
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("AddInfo_ReSelect_Employee").Click
	wait 4
	Dim counter
	counter = Environment.Value("counter")
	counter = counter + 1
	DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
	Employee_Search()
 
End Function


Function TC13_E2E_Scenario_TPAOverride()
	
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Click

End Function


