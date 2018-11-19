
Function Login()

	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Login Page "
	Systemutil.CloseProcessByName "iexplore.exe"
	Systemutil.Run "iexplore.exe", Environment.Value("CP_URL")
	Browser("ClaimsBrowser").Page("LoginPage").Sync
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("UserIdentifier").Set Environment.Value("CP_LoginId")
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("Password").Set Environment.Value("CP_LoginPassword") 
	Browser("ClaimsBrowser").Page("LoginPage").WebButton("Log In").Click
	Browser("ClaimsBrowser").Sync
End function


Function Select_Automobile()

	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Select Auto "
	Browser("name:=CCC.*").Page("title:=CCC.*").Link("html tag:=A","name:=My Group").Click
	Browser("name:=CCC.*").Page("title:=CCC.*").WebList("html id:=objWorkBasketSelect","html tag:=SELECT","name:=select").Select "Automobile"
	Browser("ClaimsBrowser").Sync
End function


Function Select_Workitem()

	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Select WorkItem "
'	Browser("Customer_Browser").Page("WorkList_Basket").WebElement("SortDate").Click
	wait(3)
	Browser("Customer_Browser").highlight
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
	Browser("ClaimsBrowser").Sync
End Function


Function Customer_Search()

	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Customer Search "
	Dim objBrwpage_CustomerSearch

	set objBrwpage_CustomerSearch=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame") 
	objBrwpage_CustomerSearch.WebButton("Customer Search").Click 
	Wait(3)
	If (DataTable("Add_NewCustomer_Flag","Common Data") = "FALSE") Then
		If objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Exist(10) Then
		   objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Set CDATE(DataTable("CS_Accident_Date","Common Data"))	
		End If		
		objBrwpage_CustomerSearch.WebEdit("CS_CustomerName").Set DataTable("CS_Customer_Name","Common Data")
		objBrwpage_CustomerSearch.WebEdit("CS_SiteCode").Set DataTable("CS_SiteCode","Common Data")
		objBrwpage_CustomerSearch.WebButton("CS_Search").Click	
		Browser("ClaimsBrowser").Sync	
		Index=1
		while index<>0'''Here the condition will waits till Web Table load
			If (objBrwpage_CustomerSearch.webelement("CS_No_Matching_Data").Exist(20) ) Then
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
			
		End If
	ElseIf DataTable("Add_NewCustomer_Flag","Common Data") = "TRUE" Then		
		Add_NewCustomer()
	Else
		Employee_Search()
	End If
	Browser("ClaimsBrowser").Sync
End Function


Function Add_NewCustomer()

		Environment.value("str_ScreenName") = "Carepoint - Auto >>>> Add New Customer "
'		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Add New Customer").Click
'		Browser("ClaimsBrowser").sync
'		wait(1)
'        Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Name").Set DataTable("AddCust_Name","Common Data")
'		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Address").Set DataTable("AddCust_Address","Common Data")
'		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Zip").Set DataTable("AddCust_Zip","Common Data")
'		wait(1)
'		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Phone").Set DataTable("AddCust_Phone","Common Data")
'		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Email").Set DataTable("AddCust_Email","Common Data")
'		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_SiteCode").Set DataTable("CS_SiteCode","Common Data")
'		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("AddCust_Submit").Click
		Browser("ClaimsBrowser").Page("CCC Bus Admin Portal 7.1").Frame("actionIFrame").WebButton("Add New Customer").Click
		Browser("ClaimsBrowser").sync
		wait(1)
        Browser("ClaimsBrowser").Page("CCC Bus Admin Portal 7.1").Frame("actionIFrame").WebEdit("AddCust_Name").Set DataTable("AddCustomer_CustomerName","Common Data")
		Browser("ClaimsBrowser").Page("CCC Bus Admin Portal 7.1").Frame("actionIFrame").WebEdit("AddCust_Address").Set DataTable("AddCustomer_Addr1","Common Data")
		Browser("ClaimsBrowser").Page("CCC Bus Admin Portal 7.1").Frame("actionIFrame").WebEdit("AddCust_Zip").Set DataTable("AddCustomer_ZIP","Common Data")
		wait(1)
		Browser("ClaimsBrowser").Page("CCC Bus Admin Portal 7.1").Frame("actionIFrame").WebEdit("AddCust_Phone").Set DataTable("AddCustomer_Phone","Common Data")
		Browser("ClaimsBrowser").Page("CCC Bus Admin Portal 7.1").Frame("actionIFrame").WebEdit("AddCust_Email").Set DataTable("AddCustomer_Email","Common Data")
		Browser("ClaimsBrowser").Page("CCC Bus Admin Portal 7.1").Frame("actionIFrame").WebButton("AddCust_Submit").Click
		Browser("ClaimsBrowser").Sync
		If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(3) then
				Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click
		Else
			
		End If
		
		If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Click
			Browser("ClaimsBrowser").Sync
		End If
	
End Function


Function Employee_Search()

	Environment.value("str_ScreenName") = "Carepoint - Auto >>>>  Employee Search "
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Employee Search").Click
	Browser("ClaimsBrowser").Sync
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Emp_CustomerName").Exist(8) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Emp_CustomerName").Select DataTable("Emp_CustomerName","Common Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Emp_Search").Click
	End If
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebRadioGroup("Emp_Result").Exist(15) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebRadioGroup("Emp_Result").Select "1"
	End If
    Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Emp_Select").Click
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Click
	End If
	Browser("ClaimsBrowser").Sync
End Function


Function Extract_SCaseId ()

	SCase_Id=""
	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebElement("SCaseId").Exist(2) Then
		SCase_Id = Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebElement("SCaseId").GetROProperty ("innertext")
		Print "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  " & SCase_Id & "  +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
	End If
	Environment.Value("SCaseId") = SCase_Id 

End Function


Function Incident()
		
	Environment.value("str_ScreenName") = "Carepoint - Auto >>>>  Incident Screen "	
	
	IN_CustomerName_Full = Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("IN_CustomerName").GetROProperty("Value")
	IN_CustomerName_Full_Len = len(IN_CustomerName_Full)
	
	If (IN_CustomerName_Full_Len > 25) Then
		IN_CustomerName = Left (IN_CustomerName_Full,25)
		IN_CustomerName_Full = IN_CustomerName
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("IN_CustomerName").Set IN_CustomerName_Full
		Wait(1)
	End If
	set objBrwpage_Incident= Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	objBrwpage_Incident.WebEdit("Reporter_Name").Set  DataTable("CO_Rep_Name","Common Data")
    objBrwpage_Incident.WebEdit("Reporter_Phone").Set  DataTable("CO_Rep_Phone","Common Data")
	objBrwpage_Incident.WebEdit("Reprter_Email").Set  DataTable("CO_Rep_Email","Common Data")
	objBrwpage_Incident.WebList("Reporter_Relationship").select  DataTable("CO_Rep_Relationship","Common Data")
	objBrwpage_Incident.WebCheckBox("ContactDetails_Checkbox").Set  "ON"
	If objBrwpage_Incident.WebEdit("In_CustDetails_ZIP").GetROProperty("width")>0 Then 
		objBrwpage_Incident.WebEdit("In_CustDetails_ZIP").Set DataTable("CS_ZipCode","Common Data")	
	End If 
	objBrwpage_Incident.WebList("AccidentState").Select DataTable("IN_AccidentState","Common Data")
	objBrwpage_Incident.WebList("AccidentTime1").Select DataTable("IN_AccidentTime1","Common Data")
	objBrwpage_Incident.WebList("AccidentTime2").Select DataTable("IN_AccidentTime2","Common Data")
	objBrwpage_Incident.WebList("AccidentTime3").Select DataTable("IN_AccidentTime3","Common Data")
	Rem added below added as Accident Description is an Edit box 
	objBrwpage_Incident.WebList("IN_Claim_SubType").Select DataTable("IN_AccDescription_ClaimSubType","Common Data")
	objBrwpage_Incident.WebList("Auto_Speciality").Select  DataTable("Auto_Speciality","Common Data")
	objBrwpage_Incident.WebList("Loss_Description").Select  DataTable("Loss_Description","Common Data")
	objBrwpage_Incident.WebEdit("In_Accident_Description").Set DataTable("IN_AccDescription","Common Data")	
	If ((DataTable("IN_AccDescription","Common Data") = "Misc-304721448") OR (DataTable("IN_AccDescription","Common Data") = "Misc-304903491")) Then
		objBrwpage_Incident.WebList("IN_Claim_SubType").Select DataTable("IN_AccDescription_ClaimSubType","Common Data")		
	End If
'	If ((DataTable("Customer_Employee_SearchFlag","Common Data") = "FALSE") OR (DataTable("Add_NewCustomer_Flag","Common Data") = "TRUE")) Then
		objBrwpage_Incident.WebEdit("IN_AccidentDate").Set DataTable("CS_Accident_Date","Common Data")
'	End If
	objBrwpage_Incident.WebButton("Policly_Next>>").Click
	Browser("ClaimsBrowser").Sync
	
End Function

Function Accident_VoidReason()
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("ACC_VoidReason").Select DataTable("ACC_VoidReason","Common Data")
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("ACC_VoidSubmit").Exist(5) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("ACC_VoidSubmit").Click
	End If
    Browser("ClaimsBrowser").Sync
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Confirm").Click
End Function

Function No_Duplicates_Found() 'If Duplicate Claim Exists

	Environment.value("str_ScreenName") = "Carepoint - Auto >>>>  Duplicate Claim "	
	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("No_Duplicates_Found").Exist(5)  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("No_Duplicates_Found").Click
		Browser("ClaimsBrowser").Sync
	Else 
		'Do Nothing
	End If
End Function

Function PolicySearch()

	Environment.value("str_ScreenName") = "Carepoint - Auto >>>>  Policy Search Screen "
	set objBrwpage_PolicySearch= Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	objBrwpage_PolicySearch.WebList("PS_Apply_PolicyNumber").Select DataTable("PS_Apply_PolicyNumber","Common Data")
	
	If (DataTable("PS_Apply_PolicyNumber","Common Data") = "No")Then
		
		If DataTable("PS_OneSeries_Indeterminate","Common Data") = "ON" Then
			objBrwpage_PolicySearch.WebList("PS_OneSeries_ZnoteCode").Select DataTable("PS_OneSeries_ZnoteCode","Common Data")
		Else
			cell_data = objBrwpage_PolicySearch.WebTable("Policy_Table").GetCellData(2,1)
			Pol_Num = objBrwpage_PolicySearch.WebTable("Policy_Table").GetCellData(2,2)
			Pol_Num=Cstr(Pol_Num)
			Datatable("Special_Policy","First Party Vehicle").Value=Pol_Num
			If cell_data = "" Then
				Set polobj = objBrwpage_PolicySearch.WebTable("Policy_Table")
				Set polobj2 = polobj.ChildItem(2,1,"WebRadioGroup",0)									
				d = polobj2.getroproperty("class")						
				If d = "Radio lvInputSelection" Then
					objBrwpage_PolicySearch.WebRadioGroup("Policy_RadioButton").Click
				End if
			End if			
			If objBrwpage_PolicySearch.WebElement("innertext:=No matching policy records found.*","innerhtml:=No matching policy records found.*").Exist Then
				objBrwpage_PolicySearch.WebButton("Policy_Retrieve").Click
				Browser("ClaimsBrowser").Sync
				If objBrwpage_PolicySearch.WebButton("Policy_Indeterminate").Exist(5) Then					
					 objBrwpage_PolicySearch.WebRadioGroup("Policy_Indeterminate").Select "Indeterminate"
				End If
			End If						
			If objBrwpage_PolicySearch.WebButton("PS_Assign1Series").Exist(5) Then
				objBrwpage_PolicySearch.WebButton("PS_Assign1Series").Click
			End If
			Browser("ClaimsBrowser").Sync
		End If
		If DataTable("PS_FourSeries_Indeterminate","Common Data") = "ON" Then
			objBrwpage_PolicySearch.WebList("PS_FourSeries_ZnoteCode").Select DataTable("PS_FourSeries_ZnoteCode","Common Data")
		Else					
			cell_data = objBrwpage_PolicySearch.WebTable("Policy_Table").GetCellData(2,1)
			If cell_data = "" Then
				Set polobj = objBrwpage_PolicySearch.WebTable("Policy_Table")
				Set polobj2 = polobj.ChildItem(2,1,"WebRadioGroup",0)							
				d = polobj2.getroproperty("class")				
				If d = "Radio lvInputSelection" Then
					objBrwpage_PolicySearch.WebRadioGroup("Policy_RadioButton").Click							
				Else
				End if
			End if
			If objBrwpage_PolicySearch.WebElement("innertext:=No matching policy records found.*","innerhtml:=No matching policy records found.*").Exist(6) Then
				objBrwpage_PolicySearch.WebButton("Policy_Retrieve").Click
				objBrwpage_PolicySearch.WebRadioGroup("Policy_Indeterminate").Select "Indeterminate"
			End If
			objBrwpage_PolicySearch.WebButton("PS_Assign4Series").Click
			objBrwpage_PolicySearch.WebRadioGroup("Policy_RadioButton").Click
			
		End If				
		objBrwpage_PolicySearch.WebButton("Policly_Next>>").Click
		Browser("ClaimsBrowser").Sync
		
	Else
		If DataTable("PS_PolicyNumber","Common Data")<> "" Then
		    objBrwpage_PolicySearch.WebEdit("PS_PolicyNumber").Set DataTable("PS_PolicyNumber","Common Data")
		    objBrwpage_PolicySearch.WebButton("Policy_Retrieve").Click
		    Wait(2)
		    Browser("ClaimsBrowser").Sync
		End If
		Cell_data = objBrwpage_PolicySearch.WebTable("Policy_Table").GetCellData(2,1)
		If cell_data = "" Then
			Set polobj = objBrwpage_PolicySearch.WebTable("Policy_Table")
			Set polobj2 = polobj.ChildItem(2,1,"WebRadioGroup",0)				
			d = polobj2.getroproperty("class")				
			If d = "Radio lvInputSelection" Then
				objBrwpage_PolicySearch.WebRadioGroup("Policy_RadioButton").Click								
			Else
			End if
		End if
		pol_flag = False
		If objBrwpage_PolicySearch.WebElement("PS_NoMatchingData").Exist  Then
			pol_flag = True
			objBrwpage_PolicySearch.WebButton("Policy_Retrieve").Click
			objBrwpage_PolicySearch.WebRadioGroup("Policy_Indeterminate").Select "Indeterminate"
		End If
		
		If(pol_flag = True) Then 
			objBrwpage_PolicySearch.WebButton("Policly_Next>>").Click
			Browser("ClaimsBrowser").Sync
		Else
			objBrwpage_PolicySearch.WebButton("Policly_Next>>").Click
			Browser("ClaimsBrowser").Sync
		End If
	End If
End Function

Function OverRide_TPA()

	Environment.value("str_ScreenName") = "Carepoint - Auto >>>>  TPA Override "
	Wait(3)
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Click
	Browser("ClaimsBrowser").Sync
	
End Function

Function Accident_Page()
	
	Environment.value("str_ScreenName") = "Carepoint - Auto >>>>  Accident Screen "
	Set objBrwPage_Accident_Page=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	Browser("ClaimsBrowser").Sync
	objBrwPage_Accident_Page.WebList("AccidentDetails_Report_Only").Select Trim(DataTable("IN_ReportOnly","Common Data")) 
	objBrwPage_Accident_Page.WebList("ACC_AccCode").Select DataTable("ACC_AccCode","Common Data")
	objBrwPage_Accident_Page.WebList("ACC_AgentLoss").Select DataTable("ACC_AgentLoss","Common Data")
	objBrwPage_Accident_Page.WebList("ACC_LossLoc").Select DataTable("ACC_LossLoc","Common Data")
	objBrwPage_Accident_Page.WebList("ACC_SiteAddress").Select DataTable("ACC_SiteAddress","Common Data")
	Accident_SiteAddr = DataTable("ACC_SiteAddress","Common Data")
	If  (Accident_SiteAddr = "No") Then
		objBrwPage_Accident_Page.WebEdit("ACC_AccAddress1").Set DataTable("ACC_AccAddress1","Common Data")
		objBrwPage_Accident_Page.WebEdit("ACC_AccAddress2").Set DataTable("ACC_AccAddress2","Common Data")
		objBrwPage_Accident_Page.WebEdit("ACC_AccZip").Set DataTable("ACC_AccZip","Common Data")
	Else
		'Do Nothing
	End If
	' POLICE
	objBrwPage_Accident_Page.WebCheckBox("ACC_Police").Set DataTable("ACC_Police","Common Data")
	objBrwPage_Accident_Page.WebCheckBox("ACC_Fire").Set DataTable("ACC_Fire","Common Data")
	objBrwPage_Accident_Page.WebCheckBox("ACC_Ambulance").Set DataTable("ACC_Ambulance","Common Data")
	objBrwPage_Accident_Page.WebCheckBox("ACC_Other").Set DataTable("ACC_Other","Common Data")
	If DataTable("ACC_Police","Common Data") = "ON" Then
		objBrwPage_Accident_Page.WebEdit("ACC_Pol_AuthName").Set DataTable("ACC_Pol_AuthName","Common Data")
		objBrwPage_Accident_Page.WebEdit("ACC_Pol_OffName").Set DataTable("ACC_Pol_OffName","Common Data")			
		objBrwPage_Accident_Page.WebEdit("ACC_Pol_OffBadge").Set DataTable("ACC_Pol_OffBadge","Common Data")
		objBrwPage_Accident_Page.WebEdit("ACC_Pol_Report").Set DataTable("ACC_Pol_Report","Common Data")			
		objBrwPage_Accident_Page.WebEdit("ACC_Pol_OffPhone").Set DataTable("ACC_Pol_OffPhone","Common Data")
		objBrwPage_Accident_Page.WebEdit("ACC_Pol_NCIC").Set DataTable("ACC_Pol_NCIC","Common Data")			
	ElseIf ((DataTable("ACC_Fire","Common Data") = "ON") OR (DataTable("ACC_Ambulance","Common Data") = "ON") OR (DataTable("ACC_Other","Common Data") = "ON")) Then
		objBrwPage_Accident_Page.WebEdit("ACC_Ambu_AuthName").Set DataTable("ACC_Ambu_AuthName","Common Data")
		If objBrwPage_Accident_Page.WebEdit("ACC_Ambu_Report").Exist(5) Then
		   objBrwPage_Accident_Page.WebEdit("ACC_Ambu_Report").Set DataTable("ACC_Ambu_Report","Common Data")	
		End If			
	End If
	
	If DataTable("ACC_VoidIncident_Flag","Common Data") = "TRUE" Then
		objBrwPage_Accident_Page.WebButton("ACC_VoidIncident").Click
		Accident_VoidReason()
	Else		
		objBrwPage_Accident_Page.WebButton("Policly_Next>>").Click
		Browser("ClaimsBrowser").Sync
	End If
		
End Function

Function Dependent()

	Environment.value("str_ScreenName") = "Carepoint - Auto >>>>  Dependent Screen "
	Sett objBrwPage_Dependent=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	objBrwPage_Dependent.WebEdit("Dependent_FirstName").Set DataTable("Dependent_FirstName","Common Data")
	objBrwPage_Dependent.WebEdit("Dependent_LastName").Set DataTable("Dependent_LastName","Common Data")
	objBrwPage_Dependent.WebEdit("Dependent_DOB").Set DataTable("Dependent_DOB","Common Data")
	objBrwPage_Dependent.WebEdit("Dependent_SSN").Set DataTable("Dependent_SSN","Common Data")
	objBrwPage_Dependent.WebEdit("Dependent_Addr1").Set DataTable("Dependent_Addr1","Common Data")
	objBrwPage_Dependent.WebEdit("Dependent_Addr2").Set DataTable("Dependent_Addr2","Common Data")
	objBrwPage_Dependent.WebEdit("Dependent_ZIP").Set DataTable("Dependent_ZIP","Common Data")
	objBrwPage_Dependent.WebEdit("Dependent_Phone1").Set DataTable("Dependent_Phone1","Common Data")
	objBrwPage_Dependent.WebEdit("Dependent_Phone2").Set DataTable("Dependent_Phone2","Common Data")
	objBrwPage_Dependent.WebEdit("Dependent_Phone3").Set DataTable("Dependent_Phone3","Common Data")
	objBrwPage_Dependent.WebEdit("Dependent_Phone4").Set DataTable("Dependent_Phone4","Common Data")
	objBrwPage_Dependent.WebList("Dependent_RelationCode").Select DataTable("Dependent_RelationCode","Common Data")
	objBrwPage_Dependent.WebButton("Dependent_Next").Click
	Browser("ClaimsBrowser").Sync

End Function

Function Claimant_First_Party_Vehicle(ByVal vehicle_owner_index)

	Set objBrwPage_PartyInfo = Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	Owner=  Datatable.value("First_Party_Owner", "Claimant Info")
	Driver=  Datatable.value("First_Party_Driver", "Claimant Info")	
	Passenger= Datatable.value("First_Party_Passenger", "Claimant Info")	
	If DataTable("First_Party_Vehicle","Claimant Info") = "Yes" Then
		If Not VarType(vehicle_owner_index) = vbString Then
			vehicle_owner_index = Cstr(vehicle_owner_index)
		End If
		If DataTable("First_Party_Owner","Claimant Info") = "Yes" Then
			If Owner="Yes" Then 
				objBrwPage_PartyInfo.WebList("Claim_Subtype").Select  DataTable("Claim_Subtype","Claimant Info")
				objBrwPage_PartyInfo.WebList("Auto_Speciality").Select  DataTable("Auto_Speciality","Claimant Info")	
				objBrwPage_PartyInfo.WebList("CL_Claim_Options").Select "1st Party Vehicle-Owner"
				If objBrwPage_PartyInfo.WebEdit("CL_Vehicle1").Exist(5) Then
					objBrwPage_PartyInfo.WebEdit("CL_Vehicle1").Set vehicle_owner_index
				End If
				If objBrwPage_PartyInfo.WebList("CL_Address1").Exist(5) Then 
					objBrwPage_PartyInfo.WebList("CL_Address1").Select DataTable("CL_First_Owner_Address","Claimant Info")
				End If					
				objBrwPage_PartyInfo.WebButton("Go").click
				If  (DataTable("CL_First_Owner_Address","Claimant Info") = "Select…")Then
					objBrwPage_PartyInfo.WebList("Party option_sameAs").Select  "Select…"
					objBrwPage_PartyInfo.WebEdit("CL_FName").Set DataTable("CL_First_Owner_FName","Claimant Info")
					objBrwPage_PartyInfo.WebEdit("CL_LName").Set DataTable("CL_First_Owner_LName","Claimant Info")
					FirstOwner_LastName_Full=DataTable("CL_First_Owner_LName","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Towing").Set DataTable("CL_Towing","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Rental").Set DataTable("CL_Rental","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
				ElseIf  (DataTable("CL_First_Owner_Address","Claimant Info") = "Site Details")Then
					objBrwPage_PartyInfo.WebList("Party option_sameAs").Select  "Site Details"
					FirstOwner_LastName_Full = objBrwPage_PartyInfo.WebEdit("CL_LName").GetROProperty("Value")
					FirstOwner_LastName_Len = len(FirstOwner_LastName_Full)
					If (FirstOwner_LastName_Len > 20) Then
						FirstOwner_LastName_Full = Left (FirstOwner_LastName_Full,20)
						objBrwPage_PartyInfo.WebEdit("CL_LName").Set FirstOwner_LastName_Full
					End If
					objBrwPage_PartyInfo.WebCheckBox("CL_Towing").Set DataTable("CL_Towing","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Rental").Set DataTable("CL_Rental","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
				End If	
				If objBrwPage_PartyInfo.WebEdit("CL_LName").GetROProperty("Value")="" Then '''thsi is to handle lname emptyness
					objBrwPage_PartyInfo.WebEdit("CL_LName").Set DataTable("CL_First_Owner_LName","Claimant Info")							
				End If								
				objBrwPage_PartyInfo.WebButton("CL_AddToList").Click
				End If
			End If
			
		If DataTable("First_Party_Driver","Claimant Info") = "Yes" Then
			vehicle_driver_index = vehicle_owner_index
			If  Driver ="Yes" Then
'				objBrwPage_PartyInfo.WebList("Party option_sameAs").Select  DataTable("Claim_SubType","Claimant Info")
'				objBrwPage_PartyInfo.WebList("Auto Speciality").Select  DataTable("Auto_Speciality","Claimant Info")
				objBrwPage_PartyInfo.WebList("Claim_Subtype").Select  DataTable("Claim_Subtype","Claimant Info")
				objBrwPage_PartyInfo.WebList("Auto_Speciality").Select  DataTable("Auto_Speciality","Claimant Info")
'				objBrwPage_PartyInfo.WebList("CL_Address1").Select DataTable("CL_First_Owner_Address","Claimant Info")                                               							
				objBrwPage_PartyInfo.WebList("CL_Claim_Options").highlight
				objBrwPage_PartyInfo.WebList("CL_Claim_Options").click
				objBrwPage_PartyInfo.WebList("CL_Claim_Options").Select "1st Party Vehicle-Driver"
				If objBrwPage_PartyInfo.WebEdit("CL_Vehicle1").Exist(5) Then
					objBrwPage_PartyInfo.WebEdit("CL_Vehicle1").Set vehicle_driver_index
				End If
				If objBrwPage_PartyInfo.WebList("CL_Address1").Exist(5) Then 
					objBrwPage_PartyInfo.WebList("CL_Address1").Select  Datatable.Value("CL_First_Driver_Address", "Claimant Info")
				End If
'				objBrwPage_PartyInfo.WebList("Auto Speciality").Select  DataTable("Auto_Speciality","Claimant Info")
				objBrwPage_PartyInfo.WebButton("Go").click
				If  (DataTable("CL_First_Driver_Address","Claimant Info") = "")Then
					objBrwPage_PartyInfo.WebEdit("CL_FName").Set DataTable("CL_First_Driver_FName","Claimant Info")
					objBrwPage_PartyInfo.WebEdit("CL_LName").Set DataTable("CL_First_Driver_LName","Claimant Info")
					FirstDriver_LastName_Full = DataTable("CL_First_Driver_LName","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Towing").Set DataTable("CL_Towing","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Rental").Set DataTable("CL_Rental","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Injured").Set DataTable("CL_Injured","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Fatality").Set DataTable("CL_Fatality","Claimant Info")
				ElseIf  (DataTable("CL_First_Driver_Address","Claimant Info") = "Owner") Then 
					objBrwPage_PartyInfo.WebEdit("CL_FName").Set DataTable("CL_First_Driver_FName","Claimant Info")
					FirstDriver_LastName_Full = objBrwPage_PartyInfo.WebEdit("CL_LName").GetROProperty("Value")
					FirstDriver_LastName_Len = len(FirstDriver_LastName_Full)
					If (FirstDriver_LastName_Len > 25) Then
						FirstDriver_LastName_Full = Left (FirstDriver_LastName_Full,25)	'												FirstDriver_LastName_Full = FirstDriver_LastName
						objBrwPage_PartyInfo.WebEdit("CL_LName").Set FirstDriver_LastName_Full
					End If
					objBrwPage_PartyInfo.WebList("Party option_sameAs").Select  "Owner" 
					objBrwPage_PartyInfo.WebCheckBox("CL_Towing").Set DataTable("CL_Towing","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Rental").Set DataTable("CL_Rental","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Injured").Set DataTable("CL_Injured","Claimant Info")
					objBrwPage_PartyInfo.WebCheckBox("CL_Fatality").Set DataTable("CL_Fatality","Claimant Info")
				End If							
			objBrwPage_PartyInfo.WebButton("CL_AddToList").Click
			End If
		End If
		
		If DataTable("First_Party_Passenger","Claimant Info") = "Yes" Then
			vehicle_passenger_index = vehicle_owner_index
			If  Passenger= "Yes"Then
				objBrwPage_PartyInfo.WebList("CL_Claim_Options").highlight
				objBrwPage_PartyInfo.WebList("CL_Claim_Options").click
				objBrwPage_PartyInfo.WebList("CL_Claim_Options").Select "1st Party Vehicle-Passenger"
				If objBrwPage_PartyInfo.WebEdit("CL_Vehicle1").Exist(5) Then
					objBrwPage_PartyInfo.WebEdit("CL_Vehicle1").Set vehicle_passenger_index
				End If
				objBrwPage_PartyInfo.WebButton("Go").click
				objBrwPage_PartyInfo.WebEdit("CL_FName").Set DataTable("CL_First_Passenger_FName","Claimant Info")
				objBrwPage_PartyInfo.WebEdit("CL_LName").Set DataTable("CL_First_Passenger_LName","Claimant Info")
				objBrwPage_PartyInfo.WebCheckBox("CL_Injured").Set DataTable("CL_Injured","Claimant Info")
				objBrwPage_PartyInfo.WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
				objBrwPage_PartyInfo.WebCheckBox("CL_Fatality").Set DataTable("CL_Fatality","Claimant Info")
				objBrwPage_PartyInfo.WebCheckBox("CL_Rental").Set DataTable("CL_Rental","Claimant Info")
				objBrwPage_PartyInfo.WebCheckBox("CL_Towing").Set DataTable("CL_Towing","Claimant Info")
				objBrwPage_PartyInfo.WebButton("CL_AddToList").Click
			End If
		End If
	End If
	
End Function



Function Claimant_Third_Party_Vehicle(ByVal vehicle_owner_index)

	Owner= DataTable.Value("Third_Party_Owner","Claimant Info")
	Driver= DataTable.Value("Third_Party_Driver","Claimant Info")
	Passenger= DataTable.Value("Third_Party_Passenger","Claimant Info")
	If DataTable("Third_Party_Vehicle","Claimant Info") = "Yes" Then
			If Not VarType(vehicle_owner_index) = vbString Then
				vehicle_owner_index = Cstr(vehicle_owner_index)
			End If
			If DataTable("Third_Party_Owner","Claimant Info") = "Yes" Then
				If  Owner= "Yes" Then
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Claim_Subtype").Select  DataTable("Claim_Subtype","Claimant Info")
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Auto_Speciality").Select  DataTable("Auto_Speciality","Claimant Info")
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("CL_Claim_Options").Select "3rd Party Vehicle-Owner"
					If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_Vehicle1").Exist(5) Then
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_Vehicle1").Set vehicle_owner_index
					End If
					If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("CL_Address1").Exist(5) Then 
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("CL_Address1").Select DataTable("CL_Third_Owner_Address","Claimant Info")
					End If
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Go").click
					If  (DataTable("CL_Third_Owner_Address","Claimant Info") = "Select...")Then
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_FName").Set DataTable("CL_Third_Owner_FName","Claimant Info")
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_LName").Set DataTable("CL_Third_Owner_LName","Claimant Info")
						ThirdOwner_LastName_Full=DataTable("CL_Third_Owner_LName","Claimant Info")
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
					ElseIf  (DataTable("CL_Third_Owner_Address","Claimant Info") = "Site Details")Then  
								ThirdOwner_LastName_Full = Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_LName").GetROProperty("Value")
								ThirdOwner_LastName_Len = len(ThirdOwner_LastName_Full)
								If (ThirdOwner_LastName_Len > 25) Then
									ThirdOwner_LastName_Full = Left (ThirdOwner_LastName_Full,25)
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_LName").Set ThirdOwner_LastName_Full
								End If
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
					End If							
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("CL_AddToList").Click
					End If
				End If
			End If
					
			If DataTable("Third_Party_Driver","Claimant Info") = "Yes" Then
					vehicle_driver_index = vehicle_owner_index
					If  Driver ="Yes" Then
'	                    Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Party option_sameAs").Select  DataTable("Claim_SubType","Claimant Info")
'						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Auto Speciality").Select  DataTable("Auto_Speciality","Claimant Info")
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Claim_Subtype").Select  DataTable("Claim_Subtype","Claimant Info")
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Auto_Speciality").Select  DataTable("Auto_Speciality","Claimant Info")
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("CL_Claim_Options").click
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("CL_Claim_Options").Select "3rd Party Vehicle-Driver"
						If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_Vehicle1").Exist(5) Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_Vehicle1").Set vehicle_driver_index
						End If
						If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("CL_Address1").Exist(5) Then 
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("CL_Address1").Select DataTable.value("CL_Third_Driver_Address","Claimant Info")
						End If
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Go").click
						If  (DataTable("CL_Third_Driver_Address","Claimant Info") = "Select...")Then
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_FName").Set DataTable("CL_Third_Driver_FName","Claimant Info")
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_LName").Set DataTable("CL_Third_Driver_LName","Claimant Info")
									ThirdDriver_LastName_Full = DataTable("CL_Third_Driver_LName","Claimant Info")
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Injured").Set DataTable("CL_Injured","Claimant Info")
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Fatality").Set DataTable("CL_Fatality","Claimant Info")
						ElseIf  (DataTable("CL_Third_Driver_Address","Claimant Info") = "Owner")Then  
									ThirdDriver_LastName_Full = Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_LName").GetROProperty("Value")
									ThirdDriver_LastName_Len = len(ThirdDriver_LastName_Full)
									If (ThirdDriver_LastName_Len > 25) Then
											ThirdDriver_LastName = Left (ThirdDriver_LastName_Full,25)
											ThirdDriver_LastName_Full = ThirdDriver_LastName
											Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_LName").Set ThirdDriver_LastName
									End If
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Injured").Set DataTable("CL_Injured","Claimant Info")
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Fatality").Set DataTable("CL_Fatality","Claimant Info")
						End If							
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("CL_AddToList").Click
						wait(2)
					End If
				End If
						
				If DataTable("Third_Party_Passenger","Claimant Info") = "Yes" Then
					vehicle_passenger_index = vehicle_owner_index		
					If  Passenger= "Yes"Then
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("CL_Claim_Options").highlight
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("CL_Claim_Options").click
						wait(2)
                    	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("CL_Claim_Options").Select "3rd Party Vehicle-Passenger"
						wait(2)
						If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_Vehicle1").Exist(5) Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_Vehicle1").Select vehicle_passenger_index
						End If
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Go").click
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_FName").Set DataTable("CL_Third_Passenger_FName","Claimant Info")
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("CL_LName").Set DataTable("CL_Third_Passenger_LName","Claimant Info")
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Injured").Set DataTable("CL_Injured","Claimant Info")
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebCheckBox("CL_Fatality").Set DataTable("CL_Fatality","Claimant Info")
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("CL_AddToList").Click
					End If
				End If

End Function




Function Claimant_Pedestrian()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If DataTable("Pedestrian","Claimant Info") = "Yes" Then
		objBrwPage_PartyInfo.WebList("Claim_Subtype").Select  DataTable("Claim_Subtype","Claimant Info")
		objBrwPage_PartyInfo.WebList("CL_Claim_Options").Select "Pedestrian"
		objBrwPage_PartyInfo.WebEdit("CL_FName").Set DataTable("CL_Ped_FName","Claimant Info")
		objBrwPage_PartyInfo.WebEdit("CL_LName").Set DataTable("CL_Ped_LName","Claimant Info")
		objBrwPage_PartyInfo.WebButton("Go").click
		objBrwPage_PartyInfo.WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
		objBrwPage_PartyInfo.WebCheckBox("CL_Fatality").Set DataTable("CL_Fatality","Claimant Info")
		objBrwPage_PartyInfo.WebButton("CL_AddToList").Click
	End If 

End Function

Function Claimant_Third_Party_Property()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If DataTable("Third_Party_Property","Claimant Info") = "Yes" Then
		objBrwPage_PartyInfo.WebList("Claim_Subtype").Select  DataTable("Claim_Subtype","Claimant Info")
		objBrwPage_PartyInfo.WebList("CL_Claim_Options").highlight
		objBrwPage_PartyInfo.WebList("CL_Claim_Options").click
		objBrwPage_PartyInfo.WebList("CL_Claim_Options").Select "3rd Party Property"
		objBrwPage_PartyInfo.WebEdit("CL_FName").Set DataTable("CL_ThirdPP_FName","Claimant Info")
		objBrwPage_PartyInfo.WebEdit("CL_LName").Set DataTable("CL_ThirdPP_LName","Claimant Info")
		objBrwPage_PartyInfo.WebButton("Go").click
		objBrwPage_PartyInfo.WebCheckBox("CL_Attorney").Set DataTable("CL_Attorney","Claimant Info")
		objBrwPage_PartyInfo.WebButton("CL_AddToList").Click
		Browser("ClaimsBrowser").Sync
	End If
        
End Function

Function ClaimantInfo_Page()   

	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Party Info Screen "	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	
	If DataTable("First_Party_Vehicle","Claimant Info") = "Yes" Then
		vehicle_owner_index = 1
		vehicle_driver_index = 1
		vehicle_passenger_index = 1
		For vehicle_owner_index=1 to DataTable("FirstOwner_Vehicle_Count","Claimant Info")
			Call Claimant_First_Party_Vehicle(vehicle_owner_index)	
		Next
	End If
	If DataTable("Third_Party_Vehicle","Claimant Info") = "Yes" Then
		vehicle_owner_index = 1
		vehicle_driver_index = 1
		vehicle_passenger_index = 1
		For vehicle_owner_index=1 to DataTable("ThirdOwner_Vehicle_Count","Claimant Info")
			Call Claimant_Third_Party_Vehicle(vehicle_owner_index)
		Next
	End If
	If DataTable("Pedestrian","Claimant Info") = "Yes" Then
		Claimant_Pedestrian()
	End If
	If DataTable("Third_Party_Property","Claimant Info") = "Yes" Then
		Claimant_Third_Party_Property()
    End If
    
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	
End Function 



Function FirstParty_OwnerData()
		
	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party Owner - Owner Data "
	Browser("ClaimsBrowser").Sync	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")	
	If  (DataTable("CL_First_Owner_Address","Claimant Info") = "Site Details")Then 
		objBrwPage_PartyInfo.WebEdit("First_OwnerData_Address1").Set DataTable("First_OwnerData_Address1","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_OwnerData_ZIP").Set DataTable("First_OwnerData_ZIP","First Party Vehicle")
	Else
		objBrwPage_PartyInfo.WebEdit("First_OwnerData_Address1").Set DataTable("First_OwnerData_Address1","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_OwnerData_ZIP").Set DataTable("First_OwnerData_ZIP","First Party Vehicle")
	 	objBrwPage_PartyInfo.WebEdit("First_OwnerData_HomePhone").Set DataTable("First_OwnerData_HomePhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_OwnerData_WorkPhone").Set DataTable("First_OwnerData_WorkPhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_OwnerData_CellPhone").Set DataTable("First_OwnerData_CellPhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_OwnerData_Fax").Set DataTable("First_OwnerData_Fax","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_OwnerData_Email").Set DataTable("First_OwnerData_Email","First Party Vehicle")
	End If
	
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click

End Function 

Function FirstParty_Owner_VehicleData()

	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party Owner - Vehicle Data "
	Browser("ClaimsBrowser").Sync	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	objBrwPage_PartyInfo.WebEdit("First_VehicleData_VIN").Set DataTable("First_VehicleData_VIN","First Party Vehicle")
	objBrwPage_PartyInfo.Image("First_VIN_Image").Click
	Wait(120)
	Browser("name:=CCC.*").Sync
	objBrwPage_PartyInfo.WebEdit("First_VehicleData_Color").Set DataTable("First_VehicleData_Color","First Party Vehicle")
	objBrwPage_PartyInfo.WebCheckBox("First_VehicleData_Tract/Trailer").Set DataTable("First_VehicleData_Tract","First Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("First_VehicleData_Plate").Set DataTable("First_VehicleData_Plate","First Party Vehicle")
	objBrwPage_PartyInfo.WebList("First_VehicleData_State").Select DataTable("First_VehicleData_State","First Party Vehicle")	
	objBrwPage_PartyInfo.WebList("First_VehicleData_Coveragetype").Select DataTable("First_VehicleData_Coveragetype","First Party Vehicle")	
	objBrwPage_PartyInfo.WebList("First_VehicleData_AdditionalInsur").Select DataTable("First_VehicleData_AdditionalInsur","First Party Vehicle")	
	objBrwPage_PartyInfo.WebList("name:=.*ppurposeOfUse","Index:=0").Select DataTable("First_VehicleData_Purpose_of_Use","First Party Vehicle")
	If DataTable("First_VehicleData_AdditionalInsur","First Party Vehicle") = "Yes" Then
		objBrwPage_PartyInfo.WebEdit("First_VData_InsurInfo_CompanyName").Set DataTable("First_VData_InsurInfo_CompanyName","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_VData_InsurInfo_Phone").Set DataTable("First_VData_InsurInfo_Phone","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_VData_InsurInfo_Policy").Set DataTable("First_VData_InsurInfo_Policy","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_VData_InsurInfo_Claim/Loss").Set DataTable("First_VData_InsurInfo_Claim","First Party Vehicle")	
	End If
	objBrwPage_PartyInfo.WebList("First_VehicleData_CurrentLoan").Select DataTable("First_VehicleData_CurrentLoan","First Party Vehicle")	
	If DataTable("First_VehicleData_CurrentLoan","First Party Vehicle") = "Yes" Then
		objBrwPage_PartyInfo.WebEdit("First_VData_LossPayInfo_LossPayName").Set DataTable("First_VData_LossPayInfo_LossPayName","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_VData_LossPayInfo_Address1").Set DataTable("First_VData_LossPayInfo_Address1","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_VData_LossPayInfo_ZIP").Set DataTable("First_VData_LossPayInfo_ZIP","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_VData_LossPayInfo_Phone").Set DataTable("First_VData_LossPayInfo_Phone","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_VData_LossPayInfo_Fax").Set DataTable("First_VData_LossPayInfo_Fax","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_VData_LossPayInfo_Email").Set DataTable("First_VData_LossPayInfo_Email","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_VData_LossPayInfo_Account").Set DataTable("First_VData_LossPayInfo_Account","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_VData_LossPayInfo_Remarks").Set DataTable("First_VData_LossPayInfo_Remarks","First Party Vehicle")
	End If
	If objBrwPage_PartyInfo.WebList("Inventory_Vehicle").Exist(5) Then  
		objBrwPage_PartyInfo.WebList("InvolvedvehicleType").Select DataTable("InvolvedvehicleType","Common Data")
		objBrwPage_PartyInfo.WebList("Inventory_Vehicle").Select DataTable("Inventory_Vehicle","Common Data")
	End If 
	
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click

End Function

Function FirstParty_Owner_Attorney()	

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("First Party Vehicle")
	If objBrwPage_PartyInfo.WebEdit("First_Att_Own_FirmName").Exist(5) Then				
		objBrwPage_PartyInfo.WebEdit("First_Att_Own_FirmName").Set DataTable("First_Att_Own_FirmName","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_Att_Own_FName").Set DataTable("First_Att_Own_FName","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_Att_Own_LastName").Set DataTable("First_Att_Own_LastName","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_Att_Own_Address1").Set DataTable("First_Att_Own_Address1","First Party Vehicle")											
		objBrwPage_PartyInfo.WebEdit("First_Att_Own_ZIP").Set DataTable("First_Att_Own_ZIP","First Party Vehicle")			
		objBrwPage_PartyInfo.WebEdit("First_Att_Own_Phone").Set DataTable("First_Att_Own_Phone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_Att_Own_AlternatePhone").Set DataTable("First_Att_Own_AlternatePhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_Att_Own_Fax").Set DataTable("First_Att_Own_Fax","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_Att_Own_Email").Set DataTable("First_Att_Own_Email","First Party Vehicle")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click	
	End If
	
End Function

Function FirstParty_DriverData()
					
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If (DataTable("CL_First_Driver_Address","Claimant Info") = "Owner") Then
		objBrwPage_PartyInfo.WebList("First_DriverData_DistributionPrefer").Select DataTable("First_DriverData_DistributionPrefer","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_DriverData_SSN").Set DataTable("First_DriverData_SSN","First Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("First_DriverData_DOB").Set DataTable("First_DriverData_DOB","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_DriverData_Gender").Select DataTable("First_DriverData_Gender","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_DriverData_MaritalStatus").Select DataTable("First_DriverData_MaritalStatus","First Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("First_DriverData_DependentCount").Set DataTable("First_DriverData_DependentCount","First Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("First_DriverData_DriverLicense").Set DataTable("First_DriverData_DriverLicense","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_DriverData_StateOfIssue").Select DataTable("First_DriverData_StateOfIssue","First Party Vehicle") 																															
		objBrwPage_PartyInfo.WebEdit("First_DriverData_Fax").Set DataTable("First_DriverData_Fax","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_DriverData_CustRelation").Set DataTable("First_DriverData_CustRelation","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_DriverData_VehicleOwnerPermission").Select DataTable("First_DriverData_VehicleOwnerPermission","First Party Vehicle") 																															
	Else
		objBrwPage_PartyInfo.WebEdit("First_DriverData_Address1").Set DataTable("First_DriverData_Address1","First Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("First_DriverData_ZIP").Set DataTable("First_DriverData_ZIP","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_DriverData_HomePhone").Set DataTable("First_DriverData_HomePhone","First Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("First_DriverData_Workphone").Set DataTable("First_DriverData_Workphone","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_DriverData_Cellphone").Set DataTable("First_DriverData_Cellphone","First Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("First_DriverData_Fax").Set DataTable("First_DriverData_Fax","First Party Vehicle") 		
		objBrwPage_PartyInfo.WebEdit("First_DriverData_Email").Set DataTable("First_DriverData_Email","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_DriverData_DistributionPrefer").Select DataTable("First_DriverData_DistributionPrefer","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_DriverData_SSN").Set DataTable("First_DriverData_SSN","First Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("First_DriverData_DOB").Set DataTable("First_DriverData_DOB","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_DriverData_Gender").Select DataTable("First_DriverData_Gender","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_DriverData_MaritalStatus").Select DataTable("First_DriverData_MaritalStatus","First Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("First_DriverData_DependentCount").Set DataTable("First_DriverData_DependentCount","First Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("First_DriverData_DriverLicense").Set DataTable("First_DriverData_DriverLicense","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_DriverData_StateOfIssue").Select DataTable("First_DriverData_StateOfIssue","First Party Vehicle") 																															
		objBrwPage_PartyInfo.WebEdit("First_DriverData_CustRelation").Set DataTable("First_DriverData_CustRelation","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_DriverData_VehicleOwnerPermission").Select DataTable("First_DriverData_VehicleOwnerPermission","First Party Vehicle") 																															
	End If
	
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click

End Function

Function FirstParty_DriverInjury()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("First Party Vehicle")
	If objBrwPage_PartyInfo.WebEdit("First_Inj_Dri_InjDesc").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("First_Inj_Dri_InjDesc").Set DataTable("First_Inj_Dri_InjDesc","First Party Vehicle") 
		If objBrwPage_PartyInfo.WebEdit("First_Driver_DeathDate").Exist(5) Then					
			objBrwPage_PartyInfo.WebEdit("First_Driver_DeathDate").Set DataTable("First_Driver_DeathDate","First Party Vehicle") 
		End If
		objBrwPage_PartyInfo.WebEdit("First_Inj_Dri_InjCause").Set DataTable("First_Inj_Dri_InjCause","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_Inj_Dri_InjNature").Select DataTable("First_Inj_Dri_InjNature","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebList("First_Inj_Dri_InjBodypart").Select DataTable("First_Inj_Dri_InjBodypart","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebList("First_Inj_Dri_InjTreatment").Select DataTable("First_Inj_Dri_InjTreatment","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_Inj_Dri_InjClmtSurgery").Select DataTable("First_Inj_Dri_InjClmtSurgery","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_Inj_Dri_InjPrevSustain").Select DataTable("First_Inj_Dri_InjPrevSustain","First Party Vehicle") 		
		objBrwPage_PartyInfo.WebList("First_Inj_Dri_InjEvacService").Select DataTable("First_Inj_Dri_InjEvacService","First Party Vehicle") 
		objBrwPage_PartyInfo.WebCheckBox("First_Inj_Dri_InjSeverity").Set DataTable("First_Inj_Dri_InjSeverity","First Party Vehicle") 		
		If( (DataTable("First_Inj_Dri_InjTreatment","First Party Vehicle") <> "NO MEDICAL TREATMENT") and (DataTable("First_Inj_Dri_InjTreatment","First Party Vehicle") <> "MINOR ON-SITE REMEDIES BY EMPLOYER MEDICAL STAFF")) Then
			FirstParty_Driver_Treatment()
		End If
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click		
	End If
	
End Function

Function FirstParty_Driver_Treatment()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("First Party Vehicle")
	If  objBrwPage_PartyInfo.WebElement("FirstDriver_Treatment_Physician").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatPhy_FirstName").Set DataTable("FirstDriver_TreatPhy_FirstName","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatPhy_Addr1").Set DataTable("FirstDriver_TreatPhy_Addr1","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatPhy_ZIP").Set DataTable("FirstDriver_TreatPhy_ZIP","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatPhy_Phone").Set DataTable("FirstDriver_TreatPhy_Phone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatPhy_Fax").Set DataTable("FirstDriver_TreatPhy_Fax","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatPhy_Email").Set DataTable("FirstDriver_TreatPhy_Email","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatHos_Name").Set DataTable("FirstDriver_TreatHos_Name","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatHos_Addr1").Set DataTable("FirstDriver_TreatHos_Addr1","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatHos_ZIP").Set DataTable("FirstDriver_TreatHos_ZIP","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatHos_Phone").Set DataTable("FirstDriver_TreatHos_Phone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatHos_Fax").Set DataTable("FirstDriver_TreatHos_Fax","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstDriver_TreatHos_Email").Set DataTable("FirstDriver_TreatHos_Email","First Party Vehicle")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If

End Function

Function FirstParty_Driver_Attorney()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("First Party Vehicle")
	If objBrwPage_PartyInfo.WebEdit("First_Att_Dri_FirmName").Exist(5) Then			
		objBrwPage_PartyInfo.WebEdit("First_Att_Dri_FirmName").Set DataTable("First_Att_Dri_FirmName","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_Att_Dri_FName").Set DataTable("First_Att_Dri_FName","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_Att_Dri_LastName").Set DataTable("First_Att_Dri_LastName","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_Att_Dri_Address1").Set DataTable("First_Att_Dri_Address1","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_Att_Dri_ZIP").Set DataTable("First_Att_Dri_ZIP","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_Att_Dri_Phone").Set DataTable("First_Att_Dri_Phone","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_Att_Dri_AlternatePhone").Set DataTable("First_Att_Dri_AlternatePhone","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_Att_Dri_Fax").Set DataTable("First_Att_Dri_Fax","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_Att_Dri_Email").Set DataTable("First_Att_Dri_Email","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebEdit("First_Att_Dri_DateNotified").Set DataTable("First_Att_Dri_DateNotified","First Party Vehicle") 	
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click		
	End If	

End Function

Function FirstParty_Passenger()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("First Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("First_Passenger_Address1").Set DataTable("First_Passenger_Address1","First Party Vehicle")  
	objBrwPage_PartyInfo.WebEdit("First_Passenger_ZIP").Set DataTable("First_Passenger_ZIP","First Party Vehicle")  
	objBrwPage_PartyInfo.WebEdit("First_Passenger_HomePhone").Set DataTable("First_Passenger_HomePhone","First Party Vehicle")  
	objBrwPage_PartyInfo.WebEdit("First_Passenger_WorkPhone").Set DataTable("First_Passenger_WorkPhone","First Party Vehicle")  
	objBrwPage_PartyInfo.WebEdit("First_Passenger_CellPhone").Set DataTable("First_Passenger_CellPhone","First Party Vehicle")  
	objBrwPage_PartyInfo.WebEdit("First_Passenger_Fax").Set DataTable("First_Passenger_Fax","First Party Vehicle")  
	objBrwPage_PartyInfo.WebEdit("First_Passenger_Email").Set DataTable("First_Passenger_Email","First Party Vehicle")
	objBrwPage_PartyInfo.WebList("First_Passenger_DistributionPrefer").Select DataTable("First_Passenger_DistributionPrefer","First Party Vehicle")  
	objBrwPage_PartyInfo.WebEdit("First_Passenger_SSN").Set DataTable("First_Passenger_SSN","First Party Vehicle")		
	objBrwPage_PartyInfo.WebEdit("First_Passenger_DOB").Set DataTable("First_Passenger_DOB","First Party Vehicle")
	objBrwPage_PartyInfo.WebList("First_Passenger_Gender").Select DataTable("First_Passenger_Gender","First Party Vehicle")  
	objBrwPage_PartyInfo.WebList("First_Passenger_MaritalStatus").Select DataTable("First_Passenger_MaritalStatus","First Party Vehicle")  
	objBrwPage_PartyInfo.WebEdit("First_Passenger_DependantCount").Set DataTable("First_Passenger_DependantCount","First Party Vehicle")								  		  
	objBrwPage_PartyInfo.WebList("First_Passenger_Language").Select DataTable("First_Passenger_Language","First Party Vehicle")  																																																																																																																																													
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	
End Function

Function FirstParty_PassengerInjury()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If objBrwPage_PartyInfo.WebEdit("First_Passenger_Inj_InjDesc").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Inj_InjDesc").Set DataTable("First_Passenger_Inj_InjDesc","First Party Vehicle")
		If objBrwPage_PartyInfo.WebEdit("First_Passenger_Inj_DeathDate").Exist(5) Then					
			objBrwPage_PartyInfo.WebEdit("First_Passenger_Inj_DeathDate").Set DataTable("First_Passenger_Inj_DeathDate","First Party Vehicle") 
		End If
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Inj_InjCause").Set DataTable("First_Passenger_Inj_InjCause","First Party Vehicle") 
		objBrwPage_PartyInfo.WebList("First_Passenger_Inj_InjNature").Select DataTable("First_Passenger_Inj_InjNature","First Party Vehicle")  
		objBrwPage_PartyInfo.WebList("First_Passenger_Inj_InjPart").Select DataTable("First_Passenger_Inj_InjPart","First Party Vehicle")  
		objBrwPage_PartyInfo.WebList("First_Passenger_Inj_Treatment").Select DataTable("First_Passenger_Inj_Treatment","First Party Vehicle")  
		objBrwPage_PartyInfo.WebList("First_Passenger_Inj_ClmtSurgery").Select DataTable("First_Passenger_Inj_ClmtSurgery","First Party Vehicle")  	
		objBrwPage_PartyInfo.WebList("First_Passenger_Inj_PrevInjury").Select DataTable("First_Passenger_Inj_PrevInjury","First Party Vehicle")  	
		objBrwPage_PartyInfo.WebList("First_Passenger_Inj_EvacService").Select DataTable("First_Passenger_Inj_EvacService","First Party Vehicle")  	
		objBrwPage_PartyInfo.WebCheckBox("First_Passenger_Inj_InjSeverity").Set DataTable("First_Passenger_Inj_InjSeverity","First Party Vehicle")  	
		If( (DataTable("First_Passenger_Inj_Treatment","First Party Vehicle") <> "NO MEDICAL TREATMENT") and (DataTable("First_Passenger_Inj_Treatment","First Party Vehicle") <> "MINOR ON-SITE REMEDIES BY EMPLOYER MEDICAL STAFF")) Then
'			objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
			FirstParty_PassengerTreatment()
		End If
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If

End Function

Function FirstParty_PassengerTreatment()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If  objBrwPage_PartyInfo.WebElement("FirstDriver_Treatment_Physician").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatPhy_FirstName").Set DataTable("FirstPassenger_TreatPhy_FirstName","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatPhy_Addr1").Set DataTable("FirstPassenger_TreatPhy_Addr1","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatPhy_ZIP").Set DataTable("FirstPassenger_TreatPhy_ZIP","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatPhy_Phone").Set DataTable("FirstPassenger_TreatPhy_Phone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatPhy_Fax").Set DataTable("FirstPassenger_TreatPhy_Fax","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatPhy_Email").Set DataTable("FirstPassenger_TreatPhy_Email","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatHos_Name").Set DataTable("FirstPassenger_TreatHos_Name","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatHos_Addr1").Set DataTable("FirstPassenger_TreatHos_Addr1","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatHos_ZIP").Set DataTable("FirstPassenger_TreatHos_ZIP","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatHos_Phone").Set DataTable("FirstPassenger_TreatHos_Phone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatHos_Fax").Set DataTable("FirstPassenger_TreatHos_Fax","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("FirstPassenger_TreatHos_Email").Set DataTable("FirstPassenger_TreatHos_Email","First Party Vehicle")
	End If

End Function

Function FirstPassenger_Attorney()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_FirmName").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_FirmName").Set DataTable("First_Passenger_Att_FirmName","First Party Vehicle")  	 
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_FName").Set DataTable("First_Passenger_Att_FName","First Party Vehicle")  	 
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_LastName").Set DataTable("First_Passenger_Att_LastName","First Party Vehicle")  
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_Address1").Set DataTable("First_Passenger_Att_Address1","First Party Vehicle")  
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_ZIP").Set DataTable("First_Passenger_Att_ZIP","First Party Vehicle")  
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_Phone").Set DataTable("First_Passenger_Att_Phone","First Party Vehicle")  
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_AltPhone").Set DataTable("First_Passenger_Att_AltPhone","First Party Vehicle")  
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_Fax").Set DataTable("First_Passenger_Att_Fax","First Party Vehicle")  
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_Email").Set DataTable("First_Passenger_Att_Email","First Party Vehicle")  
		objBrwPage_PartyInfo.WebEdit("First_Passenger_Att_DateNotified").Set DataTable("First_Passenger_Att_DateNotified","First Party Vehicle")  					
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If

End Function

Function FirstPartyVehicle()  

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	Browser("ClaimsBrowser").Sync
	
	If DataTable("First_Party_Vehicle","Claimant Info") = "Yes" Then
			
			If objBrwPage_PartyInfo.Link("First_Owner_Data").Exist(5) Then
					Call FirstParty_OwnerData()
			End If
		
			If objBrwPage_PartyInfo.Link("First_Vehicle_Data").Exist(5) Then 
					Call FirstParty_Owner_VehicleData()
				   	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("WitnessList").Exist(2) Then 
						Exit Function 
					End If
			End If
		
			If objBrwPage_PartyInfo.Link("Vehicle Location").Exist(5)  Then
					If Not Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("WitnessList").Exist(2) Then 
						Call FirstParty_Owner_Vehicle_Location()
					Else
						Exit Function 
					End If
			End If
		
			If objBrwPage_PartyInfo.Link("First_Own_Attorney").Exist(5) Then
			Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party Owner - Attorney "
					If Not Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("WitnessList").Exist(2) Then 
						Call FirstParty_Owner_Attorney()
					Else
						Exit Function 
					End If
			End If
		
			If objBrwPage_PartyInfo.Link("First_Driver_Data").Exist(5) Then
			Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party Driver -  Data "
					If Not Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("WitnessList").Exist(2) Then 
						Call FirstParty_DriverData()
					Else
						Exit Function 
					End If
			End If
		
			If objBrwPage_PartyInfo.Link("First_Injury_Driver").Exist(5) Then
			Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party Driver -  Injury Screen "
					If Not Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("WitnessList").Exist(2) Then 
						Call FirstParty_DriverInjury()
					Else
						Exit Function 
					End If
			End If
		
			If objBrwPage_PartyInfo.Link("First_Own_Attorney").Exist(5) Then
			        Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party Driver -  Attorney "
					If Not Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("WitnessList").Exist(2) Then 
						Call FirstParty_Driver_Attorney()
					Else
						Exit Function 
					End If
			End If
		
			If objBrwPage_PartyInfo.Link("First_Passenger_Data").Exist(5) Then
					Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party Passenger -  Data "
					If Not Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("WitnessList").Exist(2) Then 
						Call FirstParty_Passenger()
					Else
						Exit Function 
					End If
			End If
		
			If objBrwPage_PartyInfo.Link("First_Passenger_Injury").Exist(5) Then
					Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party Passenger -  Injury Screen "
					If Not Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("WitnessList").Exist(2) Then 
						Call FirstParty_PassengerInjury()
					Else
						Exit Function 
					End If
			End If
		
			If objBrwPage_PartyInfo.Link("First_Passenger_Attorney").Exist(5) Then
					Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party Passenger -  Attorney "
					Call FirstPassenger_Attorney()
			End If
			
	End If

End Function

Function FirstPartyVehicle2()
	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If DataTable("First_Party_Vehicle","Claimant Info") = "Yes" Then
		If (objBrwPage_PartyInfo.Link("First_V2_Owner_Data").Exist(5) and objBrwPage_PartyInfo.WebElement("First_Owner_Data").Exist(5)) Then
			Call FirstParty_V2_OwnerData()
		End If
		
		If (objBrwPage_PartyInfo.Link("First_V2_Vehicle_Data").Exist(5) and objBrwPage_PartyInfo.WebElement("First_Vehicle_Data").Exist(5)) Then
			Call FirstParty_V2_Owner_VehicleData()
		End If
		
		If (objBrwPage_PartyInfo.Link("Vehicle Loss Evaluation").Exist(5) and objBrwPage_PartyInfo.WebElement("Was vehicle damaged?*").Exist(5)) Then
			Call FirstParty_V2_Owner_VehicleLossEvaluation()
		End If
		
		If (objBrwPage_PartyInfo.Link("First_V2_Own_Attorney").Exist(5) and objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_FirmName").Exist(5)) Then
			Call FirstParty_V2_Owner_Attorney()
		End If
	End If

End Function

Function FirstParty_V2_OwnerData()
		
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If  (DataTable("CL_First_Owner_Address","Claimant Info") = "Site Details")Then 
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_Address1").Set DataTable("First_OwnerData_Address1","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_ZIP").Set DataTable("First_OwnerData_ZIP","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_HomePhone").Set DataTable("First_OwnerData_HomePhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_WorkPhone").Set DataTable("First_OwnerData_WorkPhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_CellPhone").Set DataTable("First_OwnerData_CellPhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_Email").Set DataTable("First_OwnerData_Email","First Party Vehicle")
	Else
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_Address1").Set DataTable("First_OwnerData_Address1","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_ZIP").Set DataTable("First_OwnerData_ZIP","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_HomePhone").Set DataTable("First_OwnerData_HomePhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_WorkPhone").Set DataTable("First_OwnerData_WorkPhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_CellPhone").Set DataTable("First_OwnerData_CellPhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_Fax").Set DataTable("First_OwnerData_Fax","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_OwnerData_Email").Set DataTable("First_OwnerData_Email","First Party Vehicle")
	End If
	
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click

End Function 

Function FirstParty_V2_Owner_VehicleData()
	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")			
	call  SetRow("First Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("First_V2_VehicleData_VIN").Set DataTable("First_VehicleData_VIN","First Party Vehicle")
	objBrwPage_PartyInfo.Image("First_V2_VIN_Image").Click
	Wait(120)
	Browser("name:=CCC.*").Sync
	objBrwPage_PartyInfo.WebEdit("First_V2_VehicleData_Color").Set DataTable("First_VehicleData_Color","First Party Vehicle")
	objBrwPage_PartyInfo.WebCheckBox("First_V2_VehicleData_Tract/Trailer").Set DataTable("First_VehicleData_Tract","First Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("First_V2_VehicleData_Plate").Set DataTable("First_VehicleData_Plate","First Party Vehicle")
	objBrwPage_PartyInfo.WebList("First_V2_VehicleData_State").Select DataTable("First_VehicleData_State","First Party Vehicle")	
	objBrwPage_PartyInfo.WebList("First_V2_VehicleData_Coveragetype").Select DataTable("First_VehicleData_Coveragetype","First Party Vehicle")	
	objBrwPage_PartyInfo.WebList("First_V2_VehicleData_AdditionalInsur").Select DataTable("First_VehicleData_AdditionalInsur","First Party Vehicle")	
	If DataTable("First_VehicleData_AdditionalInsur","First Party Vehicle") = "Yes" Then
		objBrwPage_PartyInfo.WebEdit("First_V2Data_InsurInfo_CompanyName").Set DataTable("First_VData_InsurInfo_CompanyName","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_V2Data_InsurInfo_Phone").Set DataTable("First_VData_InsurInfo_Phone","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_V2Data_InsurInfo_Policy").Set DataTable("First_VData_InsurInfo_Policy","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_V2Data_InsurInfo_Claim/Loss").Set DataTable("First_VData_InsurInfo_Claim","First Party Vehicle")	
	End If
	objBrwPage_PartyInfo.WebList("First_V2_VehicleData_CurrentLoan").Select DataTable("First_VehicleData_CurrentLoan","First Party Vehicle")	
	If DataTable("First_VehicleData_CurrentLoan","First Party Vehicle") = "Yes" Then
		objBrwPage_PartyInfo.WebEdit("First_V2Data_LossPayInfo_LossPayName").Set DataTable("First_VData_LossPayInfo_LossPayName","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2Data_LossPayInfo_Address1").Set DataTable("First_VData_LossPayInfo_Address1","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2Data_LossPayInfo_ZIP").Set DataTable("First_VData_LossPayInfo_ZIP","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2Data_LossPayInfo_Phone").Set DataTable("First_VData_LossPayInfo_Phone","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_V2Data_LossPayInfo_Fax").Set DataTable("First_VData_LossPayInfo_Fax","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_V2Data_LossPayInfo_Email").Set DataTable("First_VData_LossPayInfo_Email","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_V2Data_LossPayInfo_Account").Set DataTable("First_VData_LossPayInfo_Account","First Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("First_V2Data_LossPayInfo_Remarks").Set DataTable("First_VData_LossPayInfo_Remarks","First Party Vehicle")
	End If
	If  Datatable("Special_Policy"," First Party Vehicle")="302594"  Then
		If objBrwPage_PartyInfo.WebElement("DM Additional Information").Exist(5) Then
			objBrwPage_PartyInfo.WebList("DM_Vehicle Involved").Select Datatable("In_DM_Vehicle Involved","First Party Vehicle")	
			objBrwPage_PartyInfo.WebList("DM_Inventrory Vehicle").Select Datatable("In_DM_Inventrory Vehicle","First Party Vehicle")
				If Datable("In_DM_Vehicle Involved","First Party Vehicle")="Unable to confirm vehicle type at this time" Then
					objBrwPage_PartyInfo.WebEdit("DM_Vehicle Type").Set Datatable("In_DM_Vehicle Type","First Party Vehicle")	
				End If
			End if	
		End IF
		
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click

End Function

Function FirstParty_V2_Owner_VehicleDamage()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("First Party Vehicle")
	objBrwPage_PartyInfo.WebList("First_V2_VehicleDamage_Yes_No").Select DataTable("First_VehicleDamage_Yes_No","First Party Vehicle")	
	objBrwPage_PartyInfo.WebEdit("First_V2_VehicleDamage_EstimateSpeed").Set DataTable("First_VehicleDamage_EstimateSpeed","First Party Vehicle")
	If DataTable("First_VehicleDamage_Yes_No","First Party Vehicle") = "Yes" Then
		objBrwPage_PartyInfo.WebList("First_V2_VehicleDamage_LossType").Select DataTable("First_VehicleDamage_LossType","First Party Vehicle")	
		objBrwPage_PartyInfo.WebRadioGroup("First_V2Damage_Area").Select DataTable("First_VDamage_Area","First Party Vehicle")
		''*************************** If Damage Area is Front  ***********************************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Front" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_Front_Hood").Set DataTable("First_VDamage_Area_Front_Hood","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_Front_Bumper").Set DataTable("First_VDamage_Area_Front_Bumper","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_Front_WindShield").Set DataTable("First_VDamage_Area_Front_WindShield","First Party Vehicle")
		End If
		'''************************** If Damage area is Driver Front  ****************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Driver Front" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DFront_Hood").Set DataTable("First_VDamage_Area_DFront_Hood","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DFront_Fender").Set DataTable("First_VDamage_Area_DFront_Fender","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DFront_Tire").Set DataTable("First_VDamage_Area_DFront_Tire","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DFront_HeadLight").Set DataTable("First_VDamage_Area_DFront_HeadLight","First Party Vehicle")
		End If
		''''************************* If Damage Area is Driver Front Door  ************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Driver Front Door" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DFrontDoor_Window").Set DataTable("First_VDamage_Area_DFrontDoor_Window","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DFrontDoor_Door").Set DataTable("First_VDamage_Area_DFrontDoor_Door","First Party Vehicle")
		End If
		''''************************* If Damage Area is Driver Rear Door **************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Driver Rear Door" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DRearDoor_Window").Set DataTable("First_VDamage_Area_DRearDoor_Window","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassRearDoor_Door").Set DataTable("First_VDamage_Area_PassRearDoor_Door","First Party Vehicle")
		End If
		'''''********************** If Damage Area is Driver Rear  ***********************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Driver Rear" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DRear_DeckLid").Set DataTable("First_VDamage_Area_DRear_DeckLid","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DRear_Fender").Set DataTable("First_VDamage_Area_DRear_Fender","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DRear_Tire").Set DataTable("First_VDamage_Area_DRear_Tire","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_DRear_TailLight").Set DataTable("First_VDamage_Area_DRear_TailLight","First Party Vehicle")
		End If
		'''''********************* If damage Area is Passenger Front ********************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Passenger Front" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassFront_Hood").Set DataTable("First_VDamage_Area_PassFront_Hood","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassFront_Fender").Set DataTable("First_VDamage_Area_PassFront_Fender","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassFront_Tire").Set DataTable("First_VDamage_Area_PassFront_Tire","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_VDamage_Area_PassFront_HeadLight").Set DataTable("First_VDamage_Area_PassFront_HeadLight","First Party Vehicle")
		End If
		'''''**********************  If damage Area is Passenger Front Door ************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Passenger Front Door" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassFrontDoor_Window").Set DataTable("First_VDamage_Area_PassFrontDoor_Window","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassFrontDoor_Mirror").Set DataTable("First_VDamage_Area_PassFrontDoor_Mirror","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassFrontDoor_Door").Set DataTable("First_VDamage_Area_PassFrontDoor_Door","First Party Vehicle")
		End If
		''''***********************  If damage Area is Passenger Rear Door **************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Passenger Rear Door" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassRearDoor_Window").Set DataTable("First_VDamage_Area_PassRearDoor_Window","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassRearDoor_Door").Set DataTable("First_VDamage_Area_PassRearDoor_Door","First Party Vehicle")
		End If
		''''**********************  If damage Area is Passenger Rear ***********************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Passenger Rear" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassRear_DeckLid").Set DataTable("First_VDamage_Area_PassRear_DeckLid","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassRear_Fender").Set DataTable("First_VDamage_Area_PassRear_Fender","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassRear_Tire").Set DataTable("First_VDamage_Area_PassRear_Tire","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_PassRear_TailLight").Set DataTable("First_VDamage_Area_PassRear_TailLight","First Party Vehicle")
		End If
		'''''********************* If damage Area is Top *****************************************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Top" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_Top_Roof").Set DataTable("First_VDamage_Area_Top_Roof","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_Top_UnderCarriage").Set DataTable("First_VDamage_Area_Top_UnderCarriage","First Party Vehicle")
		End If
		'''**********************  If damage Area is Rear ***************************************************************************************************************************************************************************************************
		If DataTable("First_VDamage_Area","First Party Vehicle") = "Rear" Then
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_Rear_Decklid").Set DataTable("First_VDamage_Area_Rear_Decklid","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_Rear_Bumper").Set DataTable("First_VDamage_Area_Rear_Bumper","First Party Vehicle")
			objBrwPage_PartyInfo.WebCheckBox("First_V2Damage_Area_Rear_Window").Set DataTable("First_VDamage_Area_Rear_Window","First Party Vehicle")
		End If
	End If
	
	objBrwPage_PartyInfo.WebList("First_V2_VehicleDamage_PersonalProperty").Select DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle")
	If DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle") = "Yes" Then
		objBrwPage_PartyInfo.WebEdit("First_V2_VehicleDamage_PropertyDesc").Set DataTable("First_VehicleDamage_PropertyDesc","First Party Vehicle")	
	End If	
	
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click

End Function

Function FirstParty_V2_Owner_LossEvaluation()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("First Party Vehicle")
	If DataTable("First_VehicleDamage_LossType","First Party Vehicle") = "Flood" Then
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Flood_WaterReachedDashBoard").Set DataTable("First_LossEval_Flood_WaterReachedDashBoard","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Flood_EngineSubmerged").Set DataTable("First_LossEval_Flood_EngineSubmerged","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Flood_SaltWater").Set DataTable("First_LossEval_Flood_SaltWater","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Flood_WaterReachedSeats").Set DataTable("First_LossEval_Flood_WaterReachedSeats","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_NotDrivable").Set DataTable("First_LossEval_NotDrivable","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_MechFailure").Set DataTable("First_LossEval_MechFailure","First Party Vehicle")
	End If
	If DataTable("First_VehicleDamage_LossType","First Party Vehicle") = "Fire" Then
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Fire_EntireInterior").Set DataTable("First_LossEval_Fire_EntireInterior","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Fire_EntireExterior").Set DataTable("First_LossEval_Fire_EntireExterior","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Fire_EntireEngine").Set DataTable("First_LossEval_Fire_EntireEngine","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_NotDrivable").Set DataTable("First_LossEval_NotDrivable","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_MechFailure").Set DataTable("First_LossEval_MechFailure","First Party Vehicle")
	End If
	If DataTable("First_VehicleDamage_LossType","First Party Vehicle") = "Vandalism/Theft" Then
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_VandelismTheft_ComponentsDamaged").Set DataTable("First_LossEval_VandelismTheft_ComponentsDamaged","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_InteriorMissing").Set DataTable("First_LossEval_InteriorMissing","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_NotDrivable").Set DataTable("First_LossEval_NotDrivable","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_MechFailure").Set DataTable("First_LossEval_MechFailure","First Party Vehicle")
	End If
	If DataTable("First_VehicleDamage_LossType","First Party Vehicle") = "Collision/Impact with Animal" Then
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Collision_VehicleRollOver").Set DataTable("First_LossEval_Collision_VehicleRollOver","First Party Vehicle")
	    objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Collision_WindShield").Set DataTable("First_LossEval_Collision_WindShield","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Collision_DamageRearWindow").Set DataTable("First_LossEval_Collision_DamageRearWindow","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Collision_InteriorSeats_DashBoard").Set DataTable("First_LossEval_Collision_InteriorSeats_DashBoard","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Collision_Device_FreeOccupants").Set DataTable("First_LossEval_Collision_Device_FreeOccupants","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Collision_VehicleUnableToStart").Set DataTable("First_LossEval_Collision_VehicleUnableToStart","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Collision_EntireSideDamaged").Set DataTable("First_LossEval_Collision_EntireSideDamaged","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Collision_AirBagsDeployed").Set DataTable("First_LossEval_Collision_AirBagsDeployed","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Collision_FluidLeak").Set DataTable("First_LossEval_Collision_FluidLeak","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_Collision_RoofBuckled").Set DataTable("First_LossEval_Collision_RoofBuckled","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_NotDrivable").Set DataTable("First_LossEval_NotDrivable","First Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("First_V2_LossEval_MechFailure").Set DataTable("First_LossEval_MechFailure","First Party Vehicle")
	End If                        
	If objBrwPage_PartyInfo.WebButton("First_V2_LossEval_CalculateTotalLoss").Exist(5) Then
		objBrwPage_PartyInfo.WebButton("First_V2_LossEval_CalculateTotalLoss").Click
	Else
		'Do Nothing
	End If
	objBrwPage_PartyInfo.WebList("First_V2_LossEval_VehicleLoc").Select DataTable("First_LossEval_VehicleLoc","First Party Vehicle")
	
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click

End Function

Function FirstParty_V2_Owner_Attorney()	
	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("First Party Vehicle")
	If objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_FirmName").Exist(5) Then				
		objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_FirmName").Set DataTable("First_Att_Own_FirmName","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_FName").Set DataTable("First_Att_Own_FName","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_LastName").Set DataTable("First_Att_Own_LastName","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_Address1").Set DataTable("First_Att_Own_Address1","First Party Vehicle")											
		objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_ZIP").Set DataTable("First_Att_Own_ZIP","First Party Vehicle")			
		objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_Phone").Set DataTable("First_Att_Own_Phone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_AlternatePhone").Set DataTable("First_Att_Own_AlternatePhone","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_Fax").Set DataTable("First_Att_Own_Fax","First Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("First_V2_Att_Own_Email").Set DataTable("First_Att_Own_Email","First Party Vehicle")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click	
	End If
	
End Function

Function ThirdPartyVehicle() '''@@@tttt
   
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")

	If DataTable("Third_Party_Vehicle","Claimant Info") = "Yes" Then
		Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Owner -  Data "
		If objBrwPage_PartyInfo.Link("Owner Data").Exist(5) Then
			Call Third_OwnerData()
		End If
		
		If objBrwPage_PartyInfo.Link("Third_Vehicle Data").Exist(5)  Then
			Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Owner -  Vehicle Data "
			Call Third_VehicleData()
		End If
		
		If objBrwPage_PartyInfo.Link("Vehicle Loss Evaluation").Exist(5)  Then 
			Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Owner - Vehicle Loss Evaluation "
			Call ThirdParty_V2_Owner_VehicleLossEvaluation()
		End If
		
		If objBrwPage_PartyInfo.Link("Third_Owner_Attorney").Exist(5) Then
			Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Owner -  Attorney "
			call Third_OwnerAttorney()
		End If
		
		If DataTable("Third_Party_Driver","Claimant Info") = "Yes" Then
		
			If objBrwPage_PartyInfo.Link("Third_DriverData").Exist(5)  Then
				Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Driver -  Data "
				Call Third_DriverData()
			End If
			If objBrwPage_PartyInfo.Link("Third_Driver_Injury").Exist(5) Then
				Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Driver -  Injury "
				Call Third_DriverInjury()
			End If
			If  objBrwPage_PartyInfo.Link("Third_Driver_Attorney").Exist(5) Then
				Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Driver -  Attorney "
				Call Third_DriverAttorney()
			End If
		End  If
		If objBrwPage_PartyInfo.Link("Third_Passenger Data").Exist(3) and objBrwPage_PartyInfo.Link("Third_Passenger Data").Exist(3)  Then
			Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Passenger -  Data "
			Call Third_Passenger()
		End If
		If objBrwPage_PartyInfo.Link("Third_Pas_Injury").Exist(3)  and objBrwPage_PartyInfo.WebEdit("Third_Pas_Injurydesc").Exist(3)  Then
			Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Passenger -  Injury "
			Call Third_PassengerInjury()
			If (DataTable("Third_Pas_InitialTreatment","Third Party Vehicle") <> "NO MEDICAL TREATMENT" AND DataTable("Third_Pas_InitialTreatment","Third Party Vehicle") <> "MINOR ON-SITE REMEDIES BY EMPLOYER MEDICAL STAFF") Then
				Call Third_Passenger_Treatment()	
			End If	
		End If
		If objBrwPage_PartyInfo.Link("Third_Pas_Attorney").Exist(3) and objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Firm").Exist(3) Then
			Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Passenger -  Attorney "
			Call Third_PassengerAttorney()
		End If
	End If
	
End Function
''''''' ******************************  3rd Party Vehicle*******************************************************************************************************************
Function Third_OwnerData()
			
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("Claimant Info")
	call  SetRow("Third Party Vehicle")
	If objBrwPage_PartyInfo.WebEdit("Third_OwnerData_FName").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("Third_OwnerData_Fax").Set DataTable("Third_OwnerData_Fax","Third Party Vehicle")  
		If (DataTable("CL_Third_Owner_Address","Claimant Info") = "Site Details")Then  
		
			objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
		Else
			objBrwPage_PartyInfo.WebEdit("Third_OwnerData_Address1").Set DataTable("Third_OwnerData_Address1","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_OwnerData_ZIP").Set DataTable("Third_OwnerData_ZIP","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_OwnerData_HomePhone").Set DataTable("Third_OwnerData_HomePhone","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_OwnerData_WorkPhone").Set DataTable("Third_OwnerData_WorkPhone","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_OwnerData_CellPhone").Set DataTable("Third_OwnerData_CellPhone","Third Party Vehicle")
			objBrwPage_PartyInfo.WebEdit("Third_OwnerData_Fax").Set DataTable("Third_OwnerData_Fax","Third Party Vehicle")  
			objBrwPage_PartyInfo.WebEdit("Third_OwnerData_Email").Set DataTable("Third_OwnerData_Email","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebList("Third_OwnerData_DistributionPrefer").Select DataTable("Third_OwnerData_DistributionPrefer","Third Party Vehicle")  
			objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
		End If
	End If

End Function
'''''''''************************************************* Third Party Vehicle's Vehicle Data  ************************************************************************************************************************************
Function Third_VehicleData()
   
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("Third Party Vehicle")
	If objBrwPage_PartyInfo.WebList("Third_VehicleData_VehicleCustody").Exist(5)  Then 
		objBrwPage_PartyInfo.WebList("Third_VehicleData_VehicleCustody").Select DataTable("Third_VehicleData_VehicleCustody","Third Party Vehicle") 
	End If
	If objBrwPage_PartyInfo.WebEdit("Third_VehicleData_VIN").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("Third_VehicleData_VIN").Set DataTable("Third_VehicleData_VIN","Third Party Vehicle") 
		objBrwPage_PartyInfo.Image("Third_VIN_Image").Click
		Wait(120)
		Browser("name:=CCC.*").Sync
		objBrwPage_PartyInfo.WebEdit("Third_VehicleData_Plate").Set DataTable("Third_VehicleData_Plate","Third Party Vehicle") 
		objBrwPage_PartyInfo.WebList("Third_VehicleData_State").Select DataTable("Third_VehicleData_State","Third Party Vehicle") 
		objBrwPage_PartyInfo.WebList("Third_VehicleData_VehicleOwnerInfo").Select DataTable("Third_VehicleData_VehicleOwnerInfo","Third Party Vehicle") 
		If DataTable("Third_VehicleData_VehicleOwnerInfo","Third Party Vehicle") = "Yes" Then
			objBrwPage_PartyInfo.WebEdit("Third_VData_InsurInfo_CmpyName").Set DataTable("Third_VData_InsurInfo_CmpyName","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_VData_InsurInfo_Phone").Set DataTable("Third_VData_InsurInfo_Phone","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_VData_InsurInfo_Policy").Set DataTable("Third_VData_InsurInfo_Policy","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_VData_InsurInfo_Claim").Set DataTable("Third_VData_InsurInfo_Claim","Third Party Vehicle") 
		End If
		If  Datatable("Special_Policy","First Party Vehicle")="302594"  Then  
			If objBrwPage_PartyInfo.WebElement("DM Additional Information").Exist(5) Then
				objBrwPage_PartyInfo.WebList("DM_Vehicle Involved").Select Datatable("In_DM_Vehicle Involved","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebList("DM_Inventrory Vehicle").Select Datatable("In_DM_Inventrory Vehicle","Third Party Vehicle")
					If Datable("In_DM_Vehicle Involved","First Party Vehicle")="Unable to confirm vehicle type at this time" Then
						objBrwPage_PartyInfo.WebEdit("DM_Vehicle Type").Set Datatable("In_DM_Vehicle Type","Third Party Vehicle")	
					End If
				End if	
		End if
	End If
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	
End Function
''''''''''''************************************************* Third Party Vehicle's Vehicle Damage  ***********************************************************************************************************************************
Function Third_VehicleDamage()
	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("Third Party Vehicle")
	If objBrwPage_PartyInfo.WebList("Third_VehDamage_Yes_No").Exist(5) Then
		objBrwPage_PartyInfo.WebList("Third_VehDamage_Yes_No").Select DataTable("Third_VehDamage_Yes_No","Third Party Vehicle") 
		objBrwPage_PartyInfo.WebList("Third_VehDamage_EstimatedSpeed").Select DataTable("Third_VehDamage_EstimatedSpeed","Third Party Vehicle")
		If DataTable("Third_VehDamage_Yes_No","Third Party Vehicle")  = "Yes" Then
			objBrwPage_PartyInfo.WebList("Third_VehDamage_LossType").Select DataTable("Third_VehDamage_LossType","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebRadioGroup("Third_VDamage_Area").Select DataTable("Third_VDamage_Area","Third Party Vehicle") 
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Front" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_Front_Hood").Set DataTable("Third_VDamage_Front_Hood","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_Front_Bumper").Set DataTable("Third_VDamage_Front_Bumper","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_Front_WindShield").Set DataTable("Third_VDamage_Front_WindShield","Third Party Vehicle")									
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Front" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DFront_Hood").Set DataTable("Third_VDamage_DFront_Hood","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DFront_Fender").Set DataTable("Third_VDamage_DFront_Fender","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DFront_Tire").Set DataTable("Third_VDamage_DFront_Tire","Third Party Vehicle")					
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DFront_HeadLight").Set DataTable("Third_VDamage_DFront_HeadLight","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Front Door" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DFrontDoor_Window").Set DataTable("Third_VDamage_DFrontDoor_Window","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DFrontDoor_Mirror").Set DataTable("Third_VDamage_DFrontDoor_Mirror","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DFrontDoor_Door").Set DataTable("Third_VDamage_DFrontDoor_Door","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Rear Door" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DRearDoor_Window").Set DataTable("Third_VDamage_DRearDoor_Window","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DRearDoor_Door").Set DataTable("Third_VDamage_DRearDoor_Door","Third Party Vehicle")						
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Rear" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DRear_DeckLid").Set DataTable("Third_VDamage_DRear_DeckLid","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DRear_Fender").Set DataTable("Third_VDamage_DRear_Fender","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DRear_Tire").Set DataTable("Third_VDamage_DRear_Tire","Third Party Vehicle")					
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_DRear_TailLight").Set DataTable("Third_VDamage_DRear_TailLight","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Front" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassFront_Hood").Set DataTable("Third_VDamage_PassFront_Hood","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassFront_Fender").Set DataTable("Third_VDamage_PassFront_Fender","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassFront_Tire").Set DataTable("Third_VDamage_PassFront_Tire","Third Party Vehicle")					
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassFront_HeadLight").Set DataTable("Third_VDamage_PassFront_HeadLight","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Front Door" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassFrontDoor_Window").Set DataTable("Third_VDamage_PassFrontDoor_Window","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassFrontDoor_Mirror").Set DataTable("Third_VDamage_PassFrontDoor_Mirror","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassFrontDoor_Door").Set DataTable("Third_VDamage_PassFrontDoor_Door","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Rear Door" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassRearDoor_Window").Set DataTable("Third_VDamage_PassRearDoor_Window","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassRearDoor_Door").Set DataTable("Third_VDamage_PassRearDoor_Door","Third Party Vehicle")	
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Rear" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassRear_DeckLid").Set DataTable("Third_VDamage_PassRear_DeckLid","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassRear_Fender").Set DataTable("Third_VDamage_PassRear_Fender","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassRear_Tire").Set DataTable("Third_VDamage_PassRear_Tire","Third Party Vehicle")					
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_PassRear_TailLight").Set DataTable("Third_VDamage_PassRear_TailLight","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Top/Bottom" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_Top_Roof").Set DataTable("Third_VDamage_Top_Roof","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_Top_UnderCarriage").Set DataTable("Third_VDamage_Top_UnderCarriage","Third Party Vehicle")	
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Rear" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_Rear_DeckLid").Set DataTable("Third_VDamage_Rear_DeckLid","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_Rear_Bumper").Set DataTable("Third_VDamage_Rear_Bumper","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_VDamage_Rear_Window").Set DataTable("Third_VDamage_Rear_Window","Third Party Vehicle")					
			End If
		End If
		objBrwPage_PartyInfo.WebList("Third_VehDamage_PersPropDamage").Select DataTable("Third_VehDamage_PersPropDamage","Third Party Vehicle")	
		If DataTable("Third_VehDamage_PersPropDamage","Third Party Vehicle") = "Yes" Then
			objBrwPage_PartyInfo.WebEdit("Third_VehDamage_DamageDesc").Set DataTable("Third_VehDamage_DamageDesc","Third Party Vehicle")	
		End If
		
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If
End Function


Function Third_LossEvaluation()
    
    Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If DataTable("Third_VehDamage_LossType","Third Party Vehicle") = "Flood" Then
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_WaterDashBoard").Set DataTable("Third_LossEval_WaterDashBoard","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_EngSubmerged").Set DataTable("Third_LossEval_EngSubmerged","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_SaltWater").Set DataTable("Third_LossEval_SaltWater","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_WaterReachSeat").Set DataTable("Third_LossEval_WaterReachSeat","Third Party Vehicle")	
	End If
	If DataTable("Third_VehDamage_LossType","Third Party Vehicle") = "Fire" Then
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_EntInterior").Set DataTable("Third_LossEval_EntInterior","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_EntExterior").Set DataTable("Third_LossEval_EntExterior","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_EntEngine").Set DataTable("Third_LossEval_EntEngine","Third Party Vehicle")	
	End If
	If DataTable("Third_VehDamage_LossType","Third Party Vehicle") = "Vandalism/Theft" Then
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_MajorCompDamaged").Set DataTable("Third_LossEval_MajorCompDamaged","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_InteriorMissing").Set DataTable("Third_LossEval_InteriorMissing","Third Party Vehicle")	
	End If
	If DataTable("Third_VehDamage_LossType","Third Party Vehicle") = "Collision/Impact with Animal" Then
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_RollOver").Set DataTable("Third_LossEval_RollOver","Third Party Vehicle")		
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_DamageToRearWindow").Set DataTable("Third_LossEval_DamageToRearWindow","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_FreeOccupants").Set DataTable("Third_LossEval_FreeOccupants","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_EntireSide").Set DataTable("Third_LossEval_EntireSide","Third Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_FluidLeak").Set DataTable("Third_LossEval_FluidLeak","Third Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_DamageWindShield").Set DataTable("Third_LossEval_DamageWindShield","Third Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_InteriorSeats").Set DataTable("Third_LossEval_InteriorSeats","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_VehUnableToStart").Set DataTable("Third_LossEval_VehUnableToStart","Third Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_AirBagsDeployed").Set DataTable("Third_LossEval_AirBagsDeployed","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_RoofBuckled").Set DataTable("Third_LossEval_RoofBuckled","Third Party Vehicle")																																					
	End If
	If objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_NotDrivable").Exist(5) Then
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_NotDrivable").Set DataTable("Third_LossEval_NotDrivable","Third Party Vehicle")	
	End If
	If 	objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_MechFailure").Exist(5) Then
		objBrwPage_PartyInfo.WebCheckBox("Third_LossEval_MechFailure").Set DataTable("Third_LossEval_MechFailure","Third Party Vehicle")	
	End If
    If objBrwPage_PartyInfo.WebButton("Third_CalculateTotalLoss").Exist(5) Then
		objBrwPage_PartyInfo.WebButton("Third_CalculateTotalLoss").Click
	End If
	objBrwPage_PartyInfo.WebList("Third_LossEval_LocOfVehicle").Select DataTable("Third_LossEval_LocOfVehicle","Third Party Vehicle")	 
	Wait(2)
	
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	
End Function


Function Third_OwnerAttorney()
	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")	
	If objBrwPage_PartyInfo.WebEdit("Third_Att_Own_FirmName").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("Third_Att_Own_FirmName").Set DataTable("Third_Att_Own_FirmName","Third Party Vehicle")		
		objBrwPage_PartyInfo.WebEdit("Third_Att_Own_FName").Set DataTable("Third_Att_Own_FName","Third Party Vehicle")		
		objBrwPage_PartyInfo.WebEdit("Third_Att_Own_LastName").Set DataTable("Third_Att_Own_LastName","Third Party Vehicle")		
		objBrwPage_PartyInfo.WebEdit("Third_Att_Own_Address1").Set DataTable("Third_Att_Own_Address1","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("Third_Att_Own_ZIP").Set DataTable("Third_Att_Own_ZIP","Third Party Vehicle")			
		Wait(2)
		objBrwPage_PartyInfo.WebEdit("Third_Att_Own_Phone").Set DataTable("Third_Att_Own_Phone","Third Party Vehicle")			
		objBrwPage_PartyInfo.WebEdit("Third_Att_Own_Fax").Set DataTable("Third_Att_Own_Fax","Third Party Vehicle")			
		objBrwPage_PartyInfo.WebEdit("Third_Att_Own_Email").Set DataTable("Third_Att_Own_Email","Third Party Vehicle")	
		
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If
	
End Function


Function Third_DriverData()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("Claimant Info")
	call  SetRow("Third Party Vehicle")
	If  (DataTable("CL_Third_Driver_Address","Claimant Info") = "Owner")Then 
		objBrwPage_PartyInfo.WebList("Third_DriverData_DistributionPrefer").Select DataTable("Third_DriverData_DistributionPrefer","Third Party Vehicle")				
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_SSN").Set DataTable("Third_DriverData_SSN","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_DOB").Set DataTable("Third_DriverData_DOB","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebList("Third_DriverData_Gender").Select DataTable("Third_DriverData_Gender","Third Party Vehicle")				
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_DependantCount").Set DataTable("Third_DriverData_DependantCount","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_DriverData_MaritalStatus").Select DataTable("Third_DriverData_MaritalStatus","Third Party Vehicle")		
		objBrwPage_PartyInfo.WebList("Third_DriverData_Language").Select DataTable("Third_DriverData_Language","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_DriverLicense").Set DataTable("Third_DriverData_DriverLicense","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_DriverData_StateOfIssue").Select DataTable("Third_DriverData_StateOfIssue","Third Party Vehicle")			
	Else
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_Address1").Set DataTable("Third_DriverData_Address1","Third Party Vehicle")		
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_ZIP").Set DataTable("Third_DriverData_ZIP","Third Party Vehicle")		
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_HomePhone").Set DataTable("Third_DriverData_HomePhone","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_WorkPhone").Set DataTable("Third_DriverData_WorkPhone","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_CellPhone").Set DataTable("Third_DriverData_CellPhone","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_Fax").Set DataTable("Third_DriverData_Fax","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_Email").Set DataTable("Third_DriverData_Email","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebList("Third_DriverData_DistributionPrefer").Select DataTable("Third_DriverData_DistributionPrefer","Third Party Vehicle")				
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_SSN").Set DataTable("Third_DriverData_SSN","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_DOB").Set DataTable("Third_DriverData_DOB","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebList("Third_DriverData_Gender").Select DataTable("Third_DriverData_Gender","Third Party Vehicle")				
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_DependantCount").Set DataTable("Third_DriverData_DependantCount","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_DriverData_MaritalStatus").Select DataTable("Third_DriverData_MaritalStatus","Third Party Vehicle")		
		objBrwPage_PartyInfo.WebList("Third_DriverData_Language").Select DataTable("Third_DriverData_Language","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebEdit("Third_DriverData_DriverLicense").Set DataTable("Third_DriverData_DriverLicense","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_DriverData_StateOfIssue").Select DataTable("Third_DriverData_StateOfIssue","Third Party Vehicle")	
	End If	
	
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click

End Function


Function Third_DriverInjury()
	
   Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
   If objBrwPage_PartyInfo.WebEdit("Third_Inj_Driver_InjDesc").Exist(5) Then	
		objBrwPage_PartyInfo.WebEdit("Third_Inj_Driver_InjDesc").Set DataTable("Third_Inj_Driver_InjDesc","Third Party Vehicle")	
		If objBrwPage_PartyInfo.WebEdit("Third_Inj_Driver_DeathDate").Exist(5) Then
			objBrwPage_PartyInfo.WebEdit("Third_Inj_Driver_DeathDate").Set DataTable("Third_Inj_Driver_DeathDate","Third Party Vehicle")	
		End If
		objBrwPage_PartyInfo.WebEdit("Third_Inj_Driver_InjCause").Set DataTable("Third_Inj_Driver_InjCause","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebList("Third_Inj_Driver_InjNature").Select DataTable("Third_Inj_Driver_InjNature","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Inj_Driver_InjPart").Select DataTable("Third_Inj_Driver_InjPart","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Inj_Driver_InjTreatment").Select DataTable("Third_Inj_Driver_InjTreatment","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Inj_Driver_InjClmtSurgery").Select DataTable("Third_Inj_Driver_InjClmtSurgery","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Inj_Driver_PrevInjury").Select DataTable("Third_Inj_Driver_PrevInjury","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Inj_Driver_EvacService").Select DataTable("Third_Inj_Driver_EvacService","Third Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("Third_Inj_Driver_SevereInjury").Set DataTable("Third_Inj_Driver_SevereInjury","Third Party Vehicle")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If
End Function


Function Third_PhysicianHospital()
	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
    objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Physi_FName").Set DataTable("Third_InjDriver_Physi_FName","Third Party Vehicle") 
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Physi_MiddleName").Set DataTable("Third_InjDriver_Physi_MiddleName","Third Party Vehicle") 
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Physi_LastName").Set DataTable("Third_InjDriver_Physi_LastName","Third Party Vehicle") 
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Physi_Address1").Set DataTable("Third_InjDriver_Physi_Address1","Third Party Vehicle") 
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Physi_ZIP").Set DataTable("Third_InjDriver_Physi_ZIP","Third Party Vehicle") 
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Physi_Phone").Set DataTable("Third_InjDriver_Physi_Phone","Third Party Vehicle") 
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Physi_Fax").Set DataTable("Third_InjDriver_Physi_Fax","Third Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Physi_Email").Set DataTable("Third_InjDriver_Physi_Email","Third Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Hosp_Name").Set DataTable("Third_InjDriver_Hosp_Name","Third Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Hosp_Address1").Set DataTable("Third_InjDriver_Hosp_Address1","Third Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Hosp_ZIP").Set DataTable("Third_InjDriver_Hosp_ZIP","Third Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Hosp_Phone").Set DataTable("Third_InjDriver_Hosp_Phone","Third Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Hosp_Fax").Set DataTable("Third_InjDriver_Hosp_Fax","Third Party Vehicle")
	objBrwPage_PartyInfo.WebEdit("Third_InjDriver_Hosp_Email").Set DataTable("Third_InjDriver_Hosp_Email","Third Party Vehicle")
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	
End Function


Function Third_DriverAttorney()
	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_FirmName").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_FirmName").Set DataTable("Third_Att_Driver_FirmName","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_FName").Set DataTable("Third_Att_Driver_FName","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_LastName").Set DataTable("Third_Att_Driver_LastName","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_Address1").Set DataTable("Third_Att_Driver_Address1","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_ZIP").Set DataTable("Third_Att_Driver_ZIP","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_Phone").Set DataTable("Third_Att_Driver_Phone","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_AltPhone").Set DataTable("Third_Att_Driver_AltPhone","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_Fax").Set DataTable("Third_Att_Driver_Fax","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_Email").Set DataTable("Third_Att_Driver_Email","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Att_Driver_DateNotified").Set DataTable("Third_Att_Driver_DateNotified","Third Party Vehicle")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If
	
End Function


Function Third_Passenger()
   
   Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
   If objBrwPage_PartyInfo.WebEdit("Third_Pas_Add1").Exist(5) Then
 		objBrwPage_PartyInfo.WebEdit("Third_Pas_Add1").Set DataTable("Third_Pas_Add1","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Zip").Set DataTable("Third_Pas_Zip","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_HomePhone").Set DataTable("Third_Pas_HomePhone","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_WorkPhone").Set DataTable("Third_Pas_WorkPhone","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Cell").Set DataTable("Third_Pas_Cell","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Fax").Set DataTable("Third_Pas_Fax","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Email").Set DataTable("Third_Pas_Email","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_DOB").Set DataTable("Third_Pas_DOB","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Pas_DistPrefer").Select DataTable("Third_Pas_DistPrefer","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Pas_Gender").Select DataTable("Third_Pas_Gender","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Pas_MaritalStatus").Select DataTable("Third_Pas_MaritalStatus","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Pas_Language").Select DataTable("Third_Pas_Language","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_SSN").Set DataTable("Third_Pas_SSN","Third Party Vehicle")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
   End If
   
End Function


Function Third_PassengerInjury()
					
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If objBrwPage_PartyInfo.WebEdit("Third_Pas_Injurydesc").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Injurydesc").Set DataTable("Third_Pas_Injurydesc","Third Party Vehicle")
		If objBrwPage_PartyInfo.WebEdit("Third_Pas_DeathDate").Exist(5) Then
	        objBrwPage_PartyInfo.WebEdit("Third_Pas_DeathDate").Set DataTable("Third_Pas_DeathDate","Third Party Vehicle")
		End if
		objBrwPage_PartyInfo.WebEdit("Third_Pas_InjCause").Set DataTable("Third_Pas_InjCause","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Pas_NatureofInjury").Select DataTable("Third_Pas_NatureofInjury","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Pas_BodyPart").Select DataTable("Third_Pas_BodyPart","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Pas_InitialTreatment").Select DataTable("Third_Pas_InitialTreatment","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Pas_ClaimantSurgery").Select DataTable("Third_Pas_ClaimantSurgery","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Pas_PreviousInjury").Select DataTable("Third_Pas_PreviousInjury","Third Party Vehicle")
		objBrwPage_PartyInfo.WebList("Third_Pas_MedEvacuation").Select DataTable("Third_Pas_MedEvacuation","Third Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("Third_Pas_SevereInjury").Set DataTable("Third_Pas_SevereInjury","Third Party Vehicle")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If

End Function


Function Third_Passenger_Treatment()
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	 If objBrwPage_PartyInfo.Link("Third_Pas_Treatment").Exist(5) then
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Physician_Fname").Set DataTable("Third_Pas_Physician_Fname","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Physician_Lname").Set DataTable("Third_Pas_Physician_Lname","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Physician_Add1").Set DataTable("Third_Pas_Physician_Add1","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Physician_Add2").Set DataTable("Third_Pas_Physician_Add2","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Physician_Zip").Set DataTable("Third_Pas_Physician_Zip","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Physician_Phone").Set DataTable("Third_Pas_Physician_Phone","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Physician_Fax").Set DataTable("Third_Pas_Physician_Fax","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Physician_Email").Set DataTable("Third_Pas_Physician_Email","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Hosp_Name").Set DataTable("Third_Pas_Hosp_Name","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Hosp_Add1").Set DataTable("Third_Pas_Hosp_Add1","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Hosp_Add2").Set DataTable("Third_Pas_Hosp_Add2","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Hosp_Zip").Set DataTable("Third_Pas_Hosp_Zip","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Hosp_Phone").Set DataTable("Third_Pas_Hosp_Phone","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Hosp_Fax").Set DataTable("Third_Pas_Hosp_Fax","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Hosp_Email").Set DataTable("Third_Pas_Hosp_Email","Third Party Vehicle")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End if
	
End Function


Function Third_PassengerAttorney()
   
   Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
   If objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Firm").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Firm").Set DataTable("Third_Pas_Att_Firm","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Fname").Set DataTable("Third_Pas_Att_Fname","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Lname").Set DataTable("Third_Pas_Att_Lname","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Add1").Set DataTable("Third_Pas_Att_Add1","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Add2").Set DataTable("Third_Pas_Att_Add2","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Zip").Set DataTable("Third_Pas_Att_Zip","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Phone1").Set DataTable("Third_Pas_Att_Phone1","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Phone2").Set DataTable("Third_Pas_Att_Phone2","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Fax").Set DataTable("Third_Pas_Att_Fax","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_Email").Set DataTable("Third_Pas_Att_Email","Third Party Vehicle")
		objBrwPage_PartyInfo.WebEdit("Third_Pas_Att_DateNotified").Set DataTable("Third_Pas_Att_DateNotified","Third Party Vehicle")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End if
	
End Function

Function ThirdPartyVehicle2()

   Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
   If DataTable("Third_Party_Vehicle","Claimant Info") = "Yes" Then
	   If (objBrwPage_PartyInfo.Link("Third_V2_OwnerData").Exist(3) And objBrwPage_PartyInfo.WebEdit("Third_V2_OwnerData_FName").Exist(3)) Then
			Third_OwnerData2()
	   End If
	   If (objBrwPage_PartyInfo.Link("Third_V2_VehicleData").Exist(3) And objBrwPage_PartyInfo.WebEdit("Third_V2_VehicleData_VIN").Exist(3)) Then
			Third_VehicleData2()
	   End If
	   If objBrwPage_PartyInfo.Link("Vehicle Loss Evaluation").Exist(3) and objBrwPage_PartyInfo.WebList("First_VehicleDamage_Yes_No").Exist(3) Then 
			ThirdParty_V2_Owner_VehicleLossEvaluation()
	   End If
	   If objBrwPage_PartyInfo.Link("Vehicle Location").Exist(3)  and objBrwPage_PartyInfo.WebElement("Vehicle Location").Exist(3)  Then
			ThirdParty_Owner_Vehicle_Location()
	   End If
   End If
  
End Function


Function Third_OwnerData2()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	call  SetRow("Claimant Info")
	call  SetRow("Third Party Vehicle")
	If objBrwPage_PartyInfo.WebEdit("Third_V2_OwnerData_FName").Exist(5) Then
		If (DataTable("CL_Third_Owner_Address","Claimant Info") = "Site Details")Then  
			Distribution_Value = objBrwPage_PartyInfo.WebList("Third_V2_OwnerData_DistributionPrefer").GetROProperty("value")
			If Distribution_Value <> "Fax" Then
				objBrwPage_PartyInfo.WebList("Third_V2_OwnerData_DistributionPrefer").Select DataTable("Third_OwnerData_DistributionPrefer","Third Party Vehicle")  
			End If
			objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click					
		Else
			objBrwPage_PartyInfo.WebEdit("Third_V2_OwnerData_Address1").Set DataTable("Third_OwnerData_Address1","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_V2_OwnerData_ZIP").Set DataTable("Third_OwnerData_ZIP","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_V2_OwnerData_HomePhone").Set DataTable("Third_OwnerData_HomePhone","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_V2_OwnerData_WorkPhone").Set DataTable("Third_OwnerData_WorkPhone","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_V2_OwnerData_CellPhone").Set DataTable("Third_OwnerData_CellPhone","Third Party Vehicle")
			objBrwPage_PartyInfo.WebEdit("Third_V2_OwnerData_Fax").Set DataTable("Third_OwnerData_Fax","Third Party Vehicle")  
			objBrwPage_PartyInfo.WebEdit("Third_V2_OwnerData_Email").Set DataTable("Third_OwnerData_Email","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebList("Third_V2_OwnerData_DistributionPrefer").Select DataTable("Third_OwnerData_DistributionPrefer","Third Party Vehicle")  
			objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
		End If
	End If
	
End Function


Function Third_VehicleData2()
   
    Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
    call  SetRow("Third Party Vehicle")
	If objBrwPage_PartyInfo.WebEdit("Third_V2_VehicleData_VIN").Exist(5) Then
		objBrwPage_PartyInfo.WebEdit("Third_V2_VehicleData_VIN").Set DataTable("Third_VehicleData_VIN","Third Party Vehicle") 
		objBrwPage_PartyInfo.Image("Third_V2_VIN").Click
		Wait(120)
		Browser("name:=CCC.*").Sync
		objBrwPage_PartyInfo.WebEdit("Third_V2_VehicleData_Color").Set DataTable("Third_VehicleData_Color","Third Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("Third_V2_VehicleData_Plate").Set DataTable("Third_VehicleData_Plate","Third Party Vehicle") 
		objBrwPage_PartyInfo.WebList("Third_V2_VehicleData_State").Select DataTable("Third_VehicleData_State","Third Party Vehicle") 
		objBrwPage_PartyInfo.WebList("Third_V2_VehicleData_VehicleOwnerInfo").Select DataTable("Third_VehicleData_VehicleOwnerInfo","Third Party Vehicle") 
		If DataTable("Third_VehicleData_VehicleOwnerInfo","Third Party Vehicle") = "Yes" Then
			objBrwPage_PartyInfo.WebEdit("Third_V2_VData_InsurInfo_CmpyName").Set DataTable("Third_VData_InsurInfo_CmpyName","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_V2_VData_InsurInfo_Phone").Set DataTable("Third_VData_InsurInfo_Phone","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_V2_VData_InsurInfo_Policy").Set DataTable("Third_VData_InsurInfo_Policy","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebEdit("Third_V2_VData_InsurInfo_Claim").Set DataTable("Third_VData_InsurInfo_Claim","Third Party Vehicle") 
		End If
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If
	
End Function


Function Third_VehicleDamage2()

    Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
    call  SetRow("Third Party Vehicle")
	If objBrwPage_PartyInfo.WebList("Third_V2_VehDamage_Yes_No").Exist(5) Then
		objBrwPage_PartyInfo.WebList("Third_V2_VehDamage_Yes_No").Select DataTable("Third_VehDamage_Yes_No","Third Party Vehicle") 
		objBrwPage_PartyInfo.WebEdit("Third_V2_VehDamage_EstimatedSpeed").Set DataTable("Third_VehDamage_EstimatedSpeed","Third Party Vehicle")
		If DataTable("Third_VehDamage_Yes_No","Third Party Vehicle")  = "Yes" Then
			objBrwPage_PartyInfo.WebList("Third_V2_VehDamage_LossType").Select DataTable("Third_VehDamage_LossType","Third Party Vehicle") 
			objBrwPage_PartyInfo.WebRadioGroup("Third_V2_VDamage_Area").Select DataTable("Third_VDamage_Area","Third Party Vehicle") 
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Front" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_Front_Hood").Set DataTable("Third_VDamage_Front_Hood","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_Front_Bumper").Set DataTable("Third_VDamage_Front_Bumper","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_Front_WindShield").Set DataTable("Third_VDamage_Front_WindShield","Third Party Vehicle")									
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Front" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DFront_Hood").Set DataTable("Third_VDamage_DFront_Hood","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DFront_Fender").Set DataTable("Third_VDamage_DFront_Fender","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DFront_Tire").Set DataTable("Third_VDamage_DFront_Tire","Third Party Vehicle")					
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DFront_HeadLight").Set DataTable("Third_VDamage_DFront_HeadLight","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Front Door" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DFrontDoor_Window").Set DataTable("Third_VDamage_DFrontDoor_Window","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DFrontDoor_Mirror").Set DataTable("Third_VDamage_DFrontDoor_Mirror","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DFrontDoor_Door").Set DataTable("Third_VDamage_DFrontDoor_Door","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Rear Door" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DRearDoor_Window").Set DataTable("Third_VDamage_DRearDoor_Window","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DRearDoor_Door").Set DataTable("Third_VDamage_DRearDoor_Door","Third Party Vehicle")						
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Rear" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DRear_DeckLid").Set DataTable("Third_VDamage_DRear_DeckLid","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DRear_Fender").Set DataTable("Third_VDamage_DRear_Fender","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DRear_Tire").Set DataTable("Third_VDamage_DRear_Tire","Third Party Vehicle")					
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_DRear_TailLight").Set DataTable("Third_VDamage_DRear_TailLight","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Front" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassFront_Hood").Set DataTable("Third_VDamage_PassFront_Hood","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassFront_Fender").Set DataTable("Third_VDamage_PassFront_Fender","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassFront_Tire").Set DataTable("Third_VDamage_PassFront_Tire","Third Party Vehicle")					
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassFront_HeadLight").Set DataTable("Third_VDamage_PassFront_HeadLight","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Front Door" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassFrontDoor_Window").Set DataTable("Third_VDamage_PassFrontDoor_Window","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassFrontDoor_Mirror").Set DataTable("Third_VDamage_PassFrontDoor_Mirror","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassFrontDoor_Door").Set DataTable("Third_VDamage_PassFrontDoor_Door","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Rear Door" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassRearDoor_Window").Set DataTable("Third_VDamage_PassRearDoor_Window","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassRearDoor_Door").Set DataTable("Third_VDamage_PassRearDoor_Door","Third Party Vehicle")	
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Rear" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassRear_DeckLid").Set DataTable("Third_VDamage_PassRear_DeckLid","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassRear_Fender").Set DataTable("Third_VDamage_PassRear_Fender","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassRear_Tire").Set DataTable("Third_VDamage_PassRear_Tire","Third Party Vehicle")					
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_PassRear_TailLight").Set DataTable("Third_VDamage_PassRear_TailLight","Third Party Vehicle")					
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Top/Bottom" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_Top_Roof").Set DataTable("Third_VDamage_Top_Roof","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_Top_UnderCarriage").Set DataTable("Third_VDamage_Top_UnderCarriage","Third Party Vehicle")	
			End If
			If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Rear" Then
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_Rear_DeckLid").Set DataTable("Third_VDamage_Rear_DeckLid","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_Rear_Bumper").Set DataTable("Third_VDamage_Rear_Bumper","Third Party Vehicle")	
				objBrwPage_PartyInfo.WebCheckBox("Third_V2_VDamage_Rear_Window").Set DataTable("Third_VDamage_Rear_Window","Third Party Vehicle")					
			End If
		End If
		objBrwPage_PartyInfo.WebList("Third_V2_VehDamage_PersPropDamage").Select DataTable("Third_VehDamage_PersPropDamage","Third Party Vehicle")	
		If DataTable("Third_VehDamage_PersPropDamage","Third Party Vehicle") = "Yes" Then
			objBrwPage_PartyInfo.WebEdit("Third_V2_VehDamage_DamageDesc").Set DataTable("Third_VehDamage_DamageDesc","Third Party Vehicle")	
		End If
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If
	
End Function


Function Third_LossEvaluation2()

	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
    call  SetRow("Third Party Vehicle")
	If DataTable("Third_VehDamage_LossType","Third Party Vehicle") = "Flood" Then
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_WaterDashBoard").Set DataTable("Third_LossEval_WaterDashBoard","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_EngSubmerged").Set DataTable("Third_LossEval_EngSubmerged","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_SaltWater").Set DataTable("Third_LossEval_SaltWater","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_WaterReachSeat").Set DataTable("Third_LossEval_WaterReachSeat","Third Party Vehicle")	
	End If
	If DataTable("Third_VehDamage_LossType","Third Party Vehicle") = "Fire" Then
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_EntInterior").Set DataTable("Third_LossEval_EntInterior","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_EntExterior").Set DataTable("Third_LossEval_EntExterior","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_EntEngine").Set DataTable("Third_LossEval_EntEngine","Third Party Vehicle")	
	End If
	If DataTable("Third_VehDamage_LossType","Third Party Vehicle") = "Vandalism/Theft" Then
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_MajorCompDamaged").Set DataTable("Third_LossEval_MajorCompDamaged","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_InteriorMissing").Set DataTable("Third_LossEval_InteriorMissing","Third Party Vehicle")	
	End If
	If DataTable("Third_VehDamage_LossType","Third Party Vehicle") = "Collision/Impact with Animal" Then
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_RollOver").Set DataTable("Third_LossEval_RollOver","Third Party Vehicle")		
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_DamageToRearWindow").Set DataTable("Third_LossEval_DamageToRearWindow","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_FreeOccupants").Set DataTable("Third_LossEval_FreeOccupants","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_EntireSide").Set DataTable("Third_LossEval_EntireSide","Third Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_FluidLeak").Set DataTable("Third_LossEval_FluidLeak","Third Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_DamageWindShield").Set DataTable("Third_LossEval_DamageWindShield","Third Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_InteriorSeats").Set DataTable("Third_LossEval_InteriorSeats","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_VehUnableToStart").Set DataTable("Third_LossEval_VehUnableToStart","Third Party Vehicle")
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_AirBagsDeployed").Set DataTable("Third_LossEval_AirBagsDeployed","Third Party Vehicle")	
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_RoofBuckled").Set DataTable("Third_LossEval_RoofBuckled","Third Party Vehicle")																																					
	End If
	If objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_NotDrivable").Exist(5) Then
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_NotDrivable").Set DataTable("Third_LossEval_NotDrivable","Third Party Vehicle")	
	End If
	If 		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_MechFailure").Exist(5) Then
		objBrwPage_PartyInfo.WebCheckBox("Third_V2_LossEval_MechFailure").Set DataTable("Third_LossEval_MechFailure","Third Party Vehicle")	
	End If
    If objBrwPage_PartyInfo.WebButton("Third_V2_CalculateTotalLoss").Exist(5) Then
		objBrwPage_PartyInfo.WebButton("Third_V2_CalculateTotalLoss").Click
	End If
	objBrwPage_PartyInfo.WebList("Third_V2_LossEval_LocOfVehicle").Select DataTable("Third_LossEval_LocOfVehicle","Third Party Vehicle")	 
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	
End Function



Function Pedestrian_Details()
		
	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Pedestrian - Information "	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	objBrwPage_PartyInfo.WebEdit("Pedes_Address1").Set DataTable("Pedes_Address1","Pedestrian") 
	objBrwPage_PartyInfo.WebEdit("Pedes_Address2").Set DataTable("Pedes_Address2","Pedestrian") 
	objBrwPage_PartyInfo.WebEdit("Pedes_ZIP").Set DataTable("Pedes_ZIP","Pedestrian") 
	objBrwPage_PartyInfo.WebEdit("Pedes_HomePhone").Set DataTable("Pedes_HomePhone","Pedestrian")  
	objBrwPage_PartyInfo.WebEdit("Pedes_WorkPhone").Set DataTable("Pedes_WorkPhone","Pedestrian")  
	objBrwPage_PartyInfo.WebEdit("Pedes_CellPhone").Set DataTable("Pedes_CellPhone","Pedestrian")  
	objBrwPage_PartyInfo.WebEdit("Pedes_Fax").Set DataTable("Pedes_Fax","Pedestrian")  
	objBrwPage_PartyInfo.WebEdit("Pedes_Email").Set DataTable("Pedes_Email","Pedestrian")  
	objBrwPage_PartyInfo.WebList("Pedes_DistributionPrefer").Select DataTable("Pedes_DistributionPrefer","Pedestrian")	
	objBrwPage_PartyInfo.WebEdit("Pedes_SSN").Set DataTable("Pedes_SSN","Pedestrian")  
	objBrwPage_PartyInfo.WebEdit("Pedes_DOB").Set DataTable("Pedes_DOB","Pedestrian")  
	objBrwPage_PartyInfo.WebList("Pedes_Gender").Select DataTable("Pedes_Gender","Pedestrian")
	objBrwPage_PartyInfo.WebList("Pedes_MaritalStatus").Select DataTable("Pedes_MaritalStatus","Pedestrian")
	objBrwPage_PartyInfo.WebEdit("Pedes_DependantCount").Set DataTable("Pedes_DependantCount","Pedestrian") 
	objBrwPage_PartyInfo.WebList("Pedes_Language").Select DataTable("Pedes_Language","Pedestrian") 
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	If objBrwPage_PartyInfo.WebEdit("Pedes_InjuryInfo_InjuryDesc").Exist(5) Then
		Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Pedestrian - Injury Info "	
		objBrwPage_PartyInfo.WebEdit("Pedes_InjuryInfo_InjuryDesc").Set DataTable("Pedes_InjuryInfo_InjuryDesc","Pedestrian") 
		objBrwPage_PartyInfo.WebEdit("Pedes_InjuryInfo_InjuryCause").Set DataTable("Pedes_InjuryInfo_InjuryCause","Pedestrian")
		objBrwPage_PartyInfo.WebList("Pedes_InjuryInfo_InjuryNature").Select DataTable("Pedes_InjuryInfo_InjuryNature","Pedestrian")  
		objBrwPage_PartyInfo.WebList("Pedes_InjuryInfo_InjuryPart").Select DataTable("Pedes_InjuryInfo_InjuryPart","Pedestrian")  
		objBrwPage_PartyInfo.WebList("Pedes_InjuryInfo_InjuryTreatment").Select DataTable("Pedes_InjuryInfo_InjuryTreatment","Pedestrian")  
		objBrwPage_PartyInfo.WebList("Pedes_InjuryInfo_ClmtSurgery").Select DataTable("Pedes_InjuryInfo_ClmtSurgery","Pedestrian")  
		objBrwPage_PartyInfo.WebList("Pedes_InjuryInfo_PrevInjSustain").Select DataTable("Pedes_InjuryInfo_PrevInjSustain","Pedestrian")  
		objBrwPage_PartyInfo.WebList("Pedes_InjuryInfo_ClmtEvacuateService").Select DataTable("Pedes_InjuryInfo_ClmtEvacuateService","Pedestrian")  
		objBrwPage_PartyInfo.WebCheckBox("Pedes_InjuryInfo_SevereInjury").Set DataTable("Pedes_InjuryInfo_SevereInjury","Pedestrian") 
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End If
	If objBrwPage_PartyInfo.Link("Treatment").Exist(5) then
		Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Pedestrian - Injury - Physician Detials "	
		objBrwPage_PartyInfo.WebEdit("Ped_Physician_Fname").Set DataTable("Ped_Physician_Fname","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Physician_Lname").Set DataTable("Ped_Physician_Lname","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Physicain_Add1").Set DataTable("Ped_Physicain_Add1","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Physician_Add2").Set DataTable("Ped_Physician_Add2","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Physician_Zip").Set DataTable("Ped_Physician_Zip","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Physician_Phone").Set DataTable("Ped_Physician_Phone","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Physician_Fax").Set DataTable("Ped_Physician_Fax","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Physician_email").Set DataTable("Ped_Physician_email","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Hosp_Name").Set DataTable("Ped_Hosp_Name","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Hosp_Add1").Set DataTable("Ped_Hosp_Add1","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Hosp_Add2").Set DataTable("Ped_Hosp_Add2","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Hosp_Zip").Set DataTable("Ped_Hosp_Zip","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Hosp_Phone").Set DataTable("Ped_Hosp_Phone","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Hosp_Fax").Set DataTable("Ped_Hosp_Fax","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Hosp_email").Set DataTable("Ped_Hosp_email","Pedestrian")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End if
	If objBrwPage_PartyInfo.WebElement("Ped_Attorney Details").Exist(5) then
		Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Pedestrian - Attorney "	
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_Firm").Set DataTable("Ped_Attorney_Firm","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_Fname").Set DataTable("Ped_Attorney_Fname","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_Lname").Set DataTable("Ped_Attorney_Lname","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_Add1").Set DataTable("Ped_Attorney_Add1","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_Add2").Set DataTable("Ped_Attorney_Add2","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_Zip").Set DataTable("Ped_Attorney_Zip","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_Phone1").Set DataTable("Ped_Attorney_Phone1","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_Phone2").Set DataTable("Ped_Attorney_Phone2","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_Fax").Set DataTable("Ped_Attorney_Fax","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_email").Set DataTable("Ped_Attorney_email","Pedestrian")
		objBrwPage_PartyInfo.WebEdit("Ped_Attorney_DateNotified").Set DataTable("Ped_Attorney_DateNotified","Pedestrian")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End if
	
End Function


Function ThirdParty_Property_Details()
	
	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Property - Property Owner "	
	Set objBrwPage_PartyInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	
	objBrwPage_PartyInfo.WebList("PropOwner_SameAddress").Select DataTable("ThirdPPOwner_SameAddress","Third Party Property") 
	If DataTable("ThirdPPOwner_SameAddress","Third Party Property") = "No" Then
		objBrwPage_PartyInfo.WebEdit("PropOwner_Address1").Set DataTable("ThirdPPOwner_Address1","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("PropOwner_Address2").Set DataTable("ThirdPPOwner_Address2","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("PropOwner_ZIP").Set DataTable("ThirdPPOwner_ZIP","Third Party Property") 
	End If
	objBrwPage_PartyInfo.WebEdit("PropOwner_HomePhone").Set DataTable("ThirdPPOwner_HomePhone","Third Party Property") 
	objBrwPage_PartyInfo.WebEdit("PropOwner_WorkPhone").Set DataTable("ThirdPPOwner_WorkPhone","Third Party Property") 
	objBrwPage_PartyInfo.WebEdit("PropOwner_CellPhone").Set DataTable("ThirdPPOwner_CellPhone","Third Party Property") 
	objBrwPage_PartyInfo.WebEdit("PropOwner_Fax").Set DataTable("ThirdPPOwner_Fax","Third Party Property") 
	objBrwPage_PartyInfo.WebEdit("PropOwner_Email").Set DataTable("ThirdPPOwner_Email","Third Party Property") 
	objBrwPage_PartyInfo.WebList("PropOwner_DistributionPrefer").Select DataTable("ThirdPPOwner_DistributionPrefer","Third Party Property") 
	objBrwPage_PartyInfo.WebList("PropOwner_Language").Select DataTable("ThirdPPOwner_Language","Third Party Property") 
	objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	
	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Property - Property Details "
	
	If objBrwPage_PartyInfo.Link("Property Details").Exist(6) Then
		objBrwPage_PartyInfo.WebRadioGroup("name:=.*PropertyLocation","Index:=0").Select "C"
		objBrwPage_PartyInfo.WebEdit("PropDetails_PropertyDesc").Set DataTable("ThirdPPDetails_PropertyDesc","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("PropDetails_DamageDesc").Set DataTable("ThirdPPDetails_DamageDesc","Third Party Property") 
		objBrwPage_PartyInfo.WebCheckBox("PropDetails_BusinessInterup").Set DataTable("ThirdPPDetails_BusinessInterup","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("PropDetails_EstLossAmount").Set DataTable("ThirdPPDetails_EstLossAmount","Third Party Property") 
		objBrwPage_PartyInfo.WebCheckBox("PropDetails_PropertyDamage").Set DataTable("ThirdPPDetails_PropertyDamage","Third Party Property")  
		objBrwPage_PartyInfo.WebList("PropDetails_InsurInfo").Select DataTable("ThirdPPDetails_InsurInfo","Third Party Property") 
		If DataTable("ThirdPPDetails_InsurInfo","Third Party Property") = "Yes" Then
			objBrwPage_PartyInfo.WebEdit("PropDetails_InsurInfo_CmpyName").Set DataTable("ThirdPPDetails_InsurInfo_CmpyName","Third Party Property") 
			objBrwPage_PartyInfo.WebEdit("PropDetails_InsurInfo_Phone").Set DataTable("ThirdPPDetails_InsurInfo_Phone","Third Party Property") 
			objBrwPage_PartyInfo.WebEdit("PropDetails_InsurInfo_Policy").Set DataTable("ThirdPPDetails_InsurInfo_Policy","Third Party Property") 	
			objBrwPage_PartyInfo.WebEdit("PropDetails_InsurInfo_Claim").Set DataTable("ThirdPPDetails_InsurInfo_Claim","Third Party Property") 	
		End If
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End if
	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 3rd Party Property - Attroney "
	
	If (objBrwPage_PartyInfo.WebElement("ThirdPP_AttorneyDetails").Exist(5)) then
		objBrwPage_PartyInfo.WebEdit("Att_FirmName").Set DataTable("ThirdPPAtt_FirmName","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("Att_FirstName").Set DataTable("ThirdPPAtt_FirstName","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("Att_LastName").Set DataTable("ThirdPPAtt_LastName","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("Att_Address1").Set DataTable("ThirdPPAtt_Address1","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("Att_Address2").Set DataTable("ThirdPPAtt_Address2","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("Att_ZIP").Set DataTable("ThirdPPAtt_ZIP","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("Att_Phone").Set DataTable("ThirdPPAtt_Phone","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("Att_AlternatePhone").Set DataTable("ThirdPPAtt_AlternatePhone","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("Att_Fax").Set DataTable("ThirdPPAtt_Fax","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("Att_Email").Set DataTable("ThirdPPAtt_Email","Third Party Property") 
		objBrwPage_PartyInfo.WebEdit("Att_NotifiedDate").Set DataTable("ThirdPPAtt_NotifiedDate","Third Party Property")
		objBrwPage_PartyInfo.WebButton("Policly_Next>>").Click
	End if

End Function


Function Witness()
	
	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Witness Screen "
	Browser("ClaimsBrowser").Sync
	Set objBrwPage_Witness=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
   	If objBrwPage_Witness.WebList("WitnessList").Exist(5) Then 
		objBrwPage_Witness.WebList("WitnessList").Select DataTable("Witness_List","Common Data") 
	End if 
	If DataTable("Witness_List","Common Data") = "Yes" Then
		objBrwPage_Witness.WebEdit("Witness_FName").Set DataTable("Witness_FName","Common Data") 
		objBrwPage_Witness.WebEdit("Witness_LastName").Set DataTable("Witness_LastName","Common Data") 
		objBrwPage_Witness.WebEdit("Witness_Address1").Set DataTable("Witness_Address1","Common Data") 
		objBrwPage_Witness.WebEdit("Witness_Address2").Set DataTable("Witness_Address2","Common Data") 
		objBrwPage_Witness.WebEdit("Witness_ZIP").Set DataTable("Witness_ZIP","Common Data")
		objBrwPage_Witness.WebEdit("Witness_PyPhone").Set DataTable("Witness_PyPhone","Common Data")
		objBrwPage_Witness.WebEdit("Witness_Fax").Set DataTable("Witness_Fax","Common Data")
		objBrwPage_Witness.WebEdit("Witness_Email").Set DataTable("Witness_Email","Common Data")
	End If
	
	objBrwPage_Witness.WebButton("Policly_Next>>").Click
	Browser("ClaimsBrowser").Sync
	
End Function

Function ClaimPreview()	

	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Claim Preview Screen "
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Policly_Next>>").Click 
	Browser("ClaimsBrowser").Sync
	
End Function

Function AdditionalInformation()
	
	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Additional Information Screen "
	Set objBrwPage_AddInfo =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
    objBrwPage_AddInfo.WebEdit("AdditionalInfoNote").Set DataTable("AdditionalInfo_Note","Common Data")
	objBrwPage_AddInfo.WebButton("name:=Next >>").Click
	Browser("ClaimsBrowser").Sync
	
End Function

Function Assignment()
	
	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Assignment Screen "
	Set objBrwPage_Assignment = Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	If  objBrwPage_Assignment.WebElement("The Accident Code selected").Exist(2) and objBrwPage_Assignment.WebElement("Accident Code").Exist(5)  Then		
		objBrwPage_Assignment.WebList("Assignment_Accident_Code").Select "#1"
		objBrwPage_Assignment.WebButton("Run Assignment").Click
		Browser("ClaimsBrowser").Sync
	End If
	If  objBrwPage_Assignment.WebElement("If there are any updates").Exist(2) Then 
		objBrwPage_Assignment.WebButton("Run Default Assignment").Click
		Browser("ClaimsBrowser").Sync
	End if
	
	If objBrwPage_Assignment.WebButton("WC_Get_Claim_Number").Exist(10) Then
	 	objBrwPage_Assignment.WebButton("WC_Get_Claim_Number").Click
	 	Browser("ClaimsBrowser").Sync
	End If
	
	If   objBrwPage_Assignment.WebElement("Second Duplicate Search").Exist(2) and objBrwPage_Assignment.WebButton("WC_No Duplicates Found").Exist(2) Then
		 objBrwPage_Assignment.WebButton("WC_No Duplicates Found").Click	
		 Browser("ClaimsBrowser").Sync
	Else
		If objBrwPage_Assignment.WebButton("Override").GetROProperty("width")>0 Then
			objBrwPage_Assignment.WebEdit("Contact_Name").Set "Test"
			objBrwPage_Assignment.WebEdit("Contact_Phone").Set "455-577-7788"
			objBrwPage_Assignment.WebList("Override_Reason").Select "Other"
			objBrwPage_Assignment.WebEdit("Reason").Set "test"
			objBrwPage_Assignment.WebButton("Override").Click
			Browser("ClaimsBrowser").Sync
		End If	
	End If
	
	If  Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*").WebButton("name:=No Duplicates Found").Exist(5) Then 
 		Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*").WebButton("name:=No Duplicates Found").Click
 		Browser("ClaimsBrowser").Sync
	End If 		
	Call GetClaimNumber()
	
End Function


Function ReassignOffice()

	Environment.value("str_ScreenName") = "Carepoint - GL  >>>> Reassign Office Screen "
	Browser("name:=CCC.*").Sync
	Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*").WebButton("name:=Reassign Office").Click
	Browser("name:=Srchssignment").Page("title:=Srchssignment").WebEdit("name:=.*PTempAssignmentPage.*pTargetCode").Set "10NLT"
	Browser("name:=Srchssignment").Page("title:=Srchssignment").WebButton("name:=Search").Click
	Set obj = Browser("name:=Srchssignment").Page("micClass:=Page")
	Set objWebElement =  obj.webtable("column names:=Assignment;Kind;Name;Name1;Code").ChildItem(2,0,"webelement",0)
	Setting.WebPackage("ReplayType") = 2
	objWebElement.FireEvent "ondblclick",,,micLeftBtn 
	Setting.WebPackage("ReplayType") = 1 
	Browser("name:=Srchssignment").Page("title:=Srchssignment").WebButton("name:=Select").Click
	Browser("ClaimsBrowser").Sync
'	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("name:= Run Default Assignment").Exist(6) Then
'		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("name:= Run Default Assignment").Click
'	End If

	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - Assignment Screen "
	
	Set objBrwPage_Assignment = Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
		
	If objBrwPage_Assignment.WebButton("WC_Get_Claim_Number").Exist(10) Then
	 	objBrwPage_Assignment.WebButton("WC_Get_Claim_Number").Click
	 	Browser("ClaimsBrowser").Sync
	End If
			
	If  Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*").WebButton("name:=No Duplicates Found").Exist(5) Then 
 		Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*").WebButton("name:=No Duplicates Found").Click
 		Browser("ClaimsBrowser").Sync
	End If 		
	Call GetClaimNumber()

End Function

Function GetClaimNumber()

	Claim_Number= Browser("ClaimsBrowser").Page("Inbox").Frame("Review_Distribution_Frame").WebTable("ClaimNumber_Table").GetROProperty("innertext")
	Claim_Number=Trim(Claim_Number)
	Claim_Number=right(Claim_Number,24)
	If InStr(Claim_Number,"Claim") > 0 Then
		Claim_Number_1=right(Claim_Number,10)
		Claim_Number_2 = 0
	else
		Arr = Split(Claim_Number, " ")
		Claim_Number_1= Arr(1)
		Claim_Number_2 = Arr(4)
	
	End If
	Environment.Value("Claim_Number_1") = Claim_Number_1
	Environment.Value("Claim_Number_2") = Claim_Number_2
	Environment.Value("NewClaimNumber") =  Claim_Number_1 & "  " &  Claim_Number_2 & "   " & Environment.Value("SCaseId")
	Print "+++++++++++++++++++++++++++++++++++++ Claim Number is +++++++++++++++ " & Environment.Value("NewClaimNumber")  & " ++++++++++++++++++++++++++++++++++++++++++++++++++++++++ "
	
End function

Function Logout()
	
	Environment.value("str_ScreenName") = "Carepoint - Auto  >>>> Logoff Screen "
	
	Browser("name:=CC.*").Page("title:=CC.*").Image("name:=Image","image type:=Image Link","Index:=0").Click
	Browser("name:=CC.*").Page("title:=CC.*").WebElement("innertext:=Log off","html id:=ItemMiddle").Click
	SystemUtil.CloseProcessByName "iexplore.exe"
	
End Function

Function TPA_Review_Distribution()

	Set objBrwPage_Review_Distribution =Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame")
	First_Claim_Number=objBrwPage_Review_Distribution.WebElement("Claim_Number").GetROProperty("innertext")
	First_Claim_Number=right(Trim(First_Claim_Number),10)
	Environment.Value("ClaimNumber1") = First_Claim_Number
	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("actionIFrame").WebButton("Complete").Exist(5) Then
    	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("actionIFrame").WebButton("Complete").Click
   	 End If		
	
End Function


'Created By :-  Srirekha Talasila
'This will handle Distributions in Review Screen 

Function Review_Distribution()
	
		Environment.value("str_ScreenName") = "Carepoint - Auto  >>>> Review Distribution Screen "
		On Error Resume Next
		Browser("name:=CCC.*").Sync
		Call GetClaimNumber()
		If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			''Log Off	
		Else
				Set Obj_Page = Browser("name:=CCC.*").Page("title:=CCC.*")
				Set obj_ActionIFrame = Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame")
				Browser("name:=CCC.*").Sync
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
			     		CommmonValue = Left(Obj_WebList(Counter).getroproperty("name"),46)
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
			     			obj_ActionIFrame.WebCheckBox(DCheckbox1).Set "ON"
			     			
			     			Set  DCheckbox2=description.Create
							DCheckbox2("micclass").value="WebCheckBox"
							DCheckbox2("name").value= CommmonValue & "$pcoverLetterInd"
							DCheckbox2("type").value= "checkbox"
							DCheckbox2("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebCheckBox(DCheckbox2).Set "ON"
			     			
			     			Set  DCheckbox3=description.Create
							DCheckbox3("micclass").value="WebCheckBox"
							DCheckbox3("name").value= CommmonValue & "$pparListInd"
							DCheckbox3("type").value= "checkbox"
							DCheckbox3("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebCheckBox(DCheckbox3).Set "ON"
			     			
			     			Set  DCheckbox4=description.Create
							DCheckbox4("micclass").value="WebCheckBox"
							DCheckbox4("name").value= CommmonValue & "$pcustomFieldsInd"
							DCheckbox4("type").value= "checkbox"
							DCheckbox4("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebCheckBox(DCheckbox4).Set "ON"
			     			
			     			Set  DCheckbox5=description.Create
							DCheckbox5("micclass").value="WebCheckBox"
							DCheckbox5("name").value= CommmonValue & "$pnotesInd"
							DCheckbox5("type").value= "checkbox"
							DCheckbox5("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebCheckBox(DCheckbox5).Set "ON"
			     			
			     			Set  DCheckbox6=description.Create
							DCheckbox6("micclass").value="WebCheckBox"
							DCheckbox6("name").value= CommmonValue & "$poriginalDocumentInd"
							DCheckbox6("type").value= "checkbox"
							DCheckbox6("name").RegularExpression = false
							wait(2)
			     			obj_ActionIFrame.WebCheckBox(DCheckbox6).Set "ON"
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
			
	
 End Function




Function FirstParty_V2_Owner_VehicleLossEvaluation()

	If DataTable("First_VehicleDamage_Yes_No","First Party Vehicle")="No"	Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_Yes_No").Select  DataTable("First_VehicleDamage_Yes_No","First Party Vehicle")                            
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_EstimateSpeed").Select DataTable("First_VehicleDamage_EstimateSpeed","First Party Vehicle")    
		If  DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle")="Yes" Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_PersonalProperty").Select DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle")	
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Description of Personal Property").Set  "Description of Personal Property"
		End If			
		
	End IF
	If DataTable("First_VehicleDamage_Yes_No","First Party Vehicle") = "Yes" Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_Yes_No").Select  DataTable("First_VehicleDamage_Yes_No","First Party Vehicle")  
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_EstimateSpeed").Select DataTable("First_VehicleDamage_EstimateSpeed","First Party Vehicle")
	End IF                   		
	If  Datatable("First_VehicleDamage_LossType","First Party Vehicle") = "Collision/Impact with Animal"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End If
	If  Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Fire"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End If
	If  Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Flood"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End If
	If  Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Other Comprehensive"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End If
	If  Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Stolen Unrecovered"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_PersonalProperty").Select DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle")	
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Description of Personal Property").Set  "Description of Personal Property"
	End If
	If Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Vandalism/Theft"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End IF
	If Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Stolen Recovered"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End IF
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Policly_Next>>").Click

End Function

Rem newly Added function
Function  AreaOfDamage()

				'*************************** If Damage Area is Front  ***********************************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Front" Then
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle") 								
							End If
				''************************** If Damage area is Driver Front  ****************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Driver Front" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle")							
							End If
				'''************************* If Damage Area is Driver Front Door  ************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Driver Front Door" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle") 
							End If
				'''************************* If Damage Area is Driver Rear Door **************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Driver Rear Door" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle")
							End If
				''''********************** If Damage Area is Driver Rear  ***********************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Driver Rear" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle")
							End If
				''''********************* If damage Area is Passenger Front ********************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Passenger Front" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle")
							End If
				''''**********************  If damage Area is Passenger Front Door ************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Passenger Front Door" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle")
							End If
				'''***********************  If damage Area is Passenger Rear Door **************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Passenger Rear Door" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle")
							End If
				'''**********************  If damage Area is Passenger Rear ***********************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Passenger Rear" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle")
							End If
				''''********************* If damage Area is Top *****************************************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Top" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle")
							End If
				''**********************  If damage Area is Rear ***************************************************************************************************************************************************************************************************
							If DataTable("First_VDamage_Area","First Party Vehicle") = "Rear" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("First_VDamage_Area","First Party Vehicle")
							End If
			

				Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_PersonalProperty").Select DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle")	
				If DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle") = "Yes" Then

					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("personal property damaged").Select DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle")
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Description_Property").Set  Datatable("Description_Property","First Party Vehicle")	
						If  Datatable("Probable Total Loss","First Party Vehicle")="No" Then
					 Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Probable Total Loss").Select Datatable("Probable Total Loss","First Party Vehicle")
					Val1=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle is Repairable").GetROProperty("visible")
					Val2=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle is Repairable").GetROProperty("value")
					If Val1="True" and Val2="No" Then
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle is Repairable").Select "Yes"
						Val=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebElement("Probable Total Loss").GetROProperty("value")
						If Val1="No" Then 						
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle is Repairable").Select "No"
						Val1=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("personal property damaged").GetROProperty("visible")
						Val2=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("personal property damaged").GetROProperty("value")
                        	
						If Val1="True" and Val2="No" Then

					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Mileage").Select Datatable("Mileage","First Party Vehicle")
					End If
					End IF
					End IF
					End IF
					End IF

					If  Datatable("Probable Total Loss","First Party Vehicle")="No" Then
			
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Probable Total Loss").Select Datatable("Probable Total Loss","First Party Vehicle")
				Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Mileage").Select Datatable("Mileage","First Party Vehicle")
					End IF		

				Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Policly_Next>>").Click
				

End Function


Function  FirstParty_Owner_Vehicle_Location()

	
	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party Owner - Vehicle Loss Evaluation "
	Browser("ClaimsBrowser").Sync	
	If DataTable("First_VehicleDamage_Yes_No","First Party Vehicle")="No"	Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_Yes_No").Select  DataTable("First_VehicleDamage_Yes_No","First Party Vehicle")                            
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_EstimateSpeed").Select DataTable("First_VehicleDamage_EstimateSpeed","First Party Vehicle")    
		If  DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle")="Yes" Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_PersonalProperty").Select DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle")	
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Description of Personal Property").Set  "Description of Personal Property"
		End If			
	End IF
	If DataTable("First_VehicleDamage_Yes_No","First Party Vehicle") = "Yes" Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_Yes_No").Select  DataTable("First_VehicleDamage_Yes_No","First Party Vehicle")  
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_EstimateSpeed").Select DataTable("First_VehicleDamage_EstimateSpeed","First Party Vehicle")
	End IF                   		
	If  Datatable("First_VehicleDamage_LossType","First Party Vehicle") = "Collision/Impact with Animal"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End If
	If  Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Fire"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End If
	If  Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Flood"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End If
	If  Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Other Comprehensive"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End If
	If  Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Stolen Unrecovered"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("First_VehicleDamage_PersonalProperty").Select DataTable("First_VehicleDamage_PersonalProperty","First Party Vehicle")	
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Description of Personal Property").Set  "Description of Personal Property"
	End If
	If Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Vandalism/Theft"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End IF
	If Datatable("First_VehicleDamage_LossType","First Party Vehicle")="Stolen Recovered"  Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle_Loss type?").Select   Datatable("First_VehicleDamage_LossType","First Party Vehicle") 
		Call AreaOfDamage()
	End IF
	
	
	Environment.value("str_ScreenName") = "Carepoint >>>> Auto - 1st Party - Vehicle Location "
	Browser("ClaimsBrowser").Sync
	
	If Datatable("First_LossEval_VehicleLoc","First Party Vehicle")="With Insured" Then
		If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Location of Vehicle").Exist(5) Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Location of Vehicle").Select Datatable("First_LossEval_VehicleLoc","First Party Vehicle")
		End IF
	End If 
	If Datatable("First_LossEval_VehicleLoc","First Party Vehicle")="With Claimant"  then 
		If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Location of Vehicle").Exist(5) Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Location of Vehicle").Select Datatable("First_LossEval_VehicleLoc","First Party Vehicle")	
		End IF
	End IF
	
	If Datatable("First_LossEval_VehicleLoc","First Party Vehicle")="Storage Facility" Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Location of Vehicle").Select Datatable("First_LossEval_VehicleLoc","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Name").Set Datatable("Storage Facility_Name","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Address1").Set Datatable("Storage Facility_Address1","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Address2").Set  Datatable("Storage Facility_Address2","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Zip").Set Datatable("Storage Facility_Zip","First Party Vehicle")	
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_PPhone").Set Datatable("Storage Facility_PPhone","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_APhone").Set  Datatable("Storage Facility_APhone","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Email").Set  Datatable("Storage Facility_Email","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Fax").Set Datatable("Storage Facility_Fax","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Storage Facility_Permission to Release").Select Datatable("Storage Facility_Permission to Release","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Storage Facility_Tow bill").Select Datatable("Storage Facility_Tow bill","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Vehicle").Set Datatable("Storage Facility_Vehicle","First Party Vehicle")	
	End If
	
	If Datatable("First_LossEval_VehicleLoc","First Party Vehicle")="Repairer" Then		
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information_Name").Set Datatable("Repairer Information_Name","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information_Address1").Set Datatable("Repairer Information_Address1","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information _Address2").Set Datatable("Repairer Information _Address2","First Party Vehicle")		
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information_Zip").Set Datatable("Repairer Information_Zip","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information_PPhone").Set Datatable("Repairer Information_PPhone","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information_Email").Set Datatable("Repairer Information_Email","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information _Fax").Set Datatable("Repairer Information _Fax","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Storage Facility_Permission to Release").Select Datatable("Storage Facility_Permission to Release","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Storage Facility_Tow bill").Select Datatable("Storage Facility_Tow bill","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Vehicle").Set Datatable("SStorage Facility_Vehicle","First Party Vehicle")
	End IF
	
	If Datatable("First_LossEval_VehicleLoc","First Party Vehicle")="Insured Other Than Site" Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Name").Set Datatable("Insured_Other_Than_Site_Name","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Add1").Set Datatable("Insured_Other_Than_Site_Add1","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Add2").Set Datatable("Insured_Other_Than_Site_Add2","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_PPhone").Set Datatable("Insured_Other_Than_Site_Pphone","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Aphone").Set Datatable("Insured_Other_Than_Site_Aphone","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Zip").Set Datatable("Insured_Other_Than_Site_Zip","First Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Email").Set  Datatable("Insured_Other_Than_Site_Email","First Party Vehicle")
	End IF
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Policly_Next>>").Click
	Browser("ClaimsBrowser").Sync
	
End Function 


Rem Function written for the third pary vehicle
Function ThirdParty_V2_Owner_VehicleLossEvaluation()
	 	If DataTable("Third_VehDamage_Yes_No","Third Party Vehicle")="No"	Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_Yes_No").Select  DataTable("Third_VehDamage_Yes_No","Third Party Vehicle")  
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_EstimatedSpeed").Select DataTable("Third_VehDamage_EstimatedSpeed","Third Party Vehicle")    
			If  DataTable("Third_VehicleDamage_PersonalProperty","Third Party Vehicle")="Yes" Then
				Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_PersPropDamage").Select DataTable("Third_VehicleDamage_PersonalProperty","Third Party Vehicle")	
			End If			
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Third_LossEval_DamageDesc").Set  "Description of Personal Property"
	    End IF
		If DataTable("Third_VehDamage_Yes_No","Third Party Vehicle") = "Yes" Then
		    Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_Yes_No").Select  DataTable("Third_VehDamage_Yes_No","Third Party Vehicle")  
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_EstimatedSpeed").Select DataTable("Third_VehDamage_EstimatedSpeed","Third Party Vehicle")
		End IF                   		
		If  Datatable("Third_VehDamage_LossType","Third Party Vehicle") = "Collision/Impact with Animal"  Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_LossType").Select   Datatable("Third_VehDamage_LossType","Third Party Vehicle") 
			Call AreaOfDamage()
		 End If	
		 If  Datatable("Third_VehDamage_LossType","Third Party Vehicle")="Fire"  Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_LossType").Select   Datatable("Third_VehDamage_LossType","Third Party Vehicle") 
			Call AreaOfDamage()
		 End If
        If  Datatable("Third_VehDamage_LossType","Third Party Vehicle")="Flood"  Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_LossType").Select   Datatable("Third_VehDamage_LossType","Third Party Vehicle") 
			Call AreaOfDamage()
		 End If
        If  Datatable("Third_VehDamage_LossType","Third Party Vehicle")="Other Comprehensive"  Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_LossType").Select   Datatable("Third_VehDamage_LossType","Third Party Vehicle") 
			Call AreaOfDamage()
		 End If
	    If  Datatable("Third_VehDamage_LossType","Third Party Vehicle")="Stolen Unrecovered"  Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_PersPropDamage").Select DataTable("Third_VehicleDamage_PersonalProperty","Third Party Vehicle")	
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Third_LossEval_DamageDesc").Set  "Description of Personal Property"
		 End If
		If Datatable("Third_VehDamage_LossType","Third Party Vehicle")="Vandalism/Theft"  Then		
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_LossType").Select   Datatable("Third_VehDamage_LossType","Third Party Vehicle") 
			Call AreaOfDamage()
		End IF
		If Datatable("Third_VehDamage_LossType","Third Party Vehicle")="Stolen Recovered"  Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_LossType").Select   Datatable("Third_VehDamage_LossType","Third Party Vehicle") 
			Call AreaOfDamage()
		End IF
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Policly_Next>>").Click
		
End Function

Rem newly Added function
Function  AreaOfDamage()

				'*************************** If Damage Area is Front  ***********************************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Front" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle") 								
							End If
				''************************** If Damage area is Driver Front  ****************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Front" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle")							
							End If
				'''************************* If Damage Area is Driver Front Door  ************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Front Door" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle") 
							End If
				'''************************* If Damage Area is Driver Rear Door **************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Rear Door" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle")
							End If
				''''********************** If Damage Area is Driver Rear  ***********************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Driver Rear" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle")
							End If
				''''********************* If damage Area is Passenger Front ********************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Front" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle")
							End If
				''''**********************  If damage Area is Passenger Front Door ************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Front Door" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle")
							End If
				'''***********************  If damage Area is Passenger Rear Door **************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Rear Door" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle")
							End If
				'''**********************  If damage Area is Passenger Rear ***********************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Passenger Rear" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle")
							End If
				''''********************* If damage Area is Top *****************************************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Top" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle")
							End If
				''**********************  If damage Area is Rear ***************************************************************************************************************************************************************************************************
							If DataTable("Third_VDamage_Area","Third Party Vehicle") = "Rear" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebRadioGroup("Front").Select  DataTable("Third_VDamage_Area","Third Party Vehicle")
							End If
			
					Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Third_VehDamage_PersPropDamage").Select DataTable("Third_VehicleDamage_PersonalProperty","Third Party Vehicle")	
					If DataTable("Third_VehicleDamage_PersonalProperty","Third Party Vehicle") = "Yes" Then
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("personal property damaged").Select DataTable("Third_VehicleDamage_PersonalProperty","Third Party Vehicle")
						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Description_Property").Set  Datatable("Description_Property","Third Party Vehicle")	
						If  Datatable("Probable_Total_Loss","Third Party Vehicle")="No" Then
							 Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Probable Total Loss").Select Datatable("Probable_Total_Loss","Third Party Vehicle")
							Val1=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle is Repairable").GetROProperty("visible")
							Val2=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle is Repairable").GetROProperty("value")
							If Val1="True" and Val2="No" Then
								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle is Repairable").Select "Yes"
								Val=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebElement("Probable Total Loss").GetROProperty("value")
								If Val1="No" Then 						
									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Vehicle is Repairable").Select "No"
									Val1=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("personal property damaged").GetROProperty("visible")
									Val2=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("personal property damaged").GetROProperty("value")
									If Val1="True" and Val2="No" Then	
										Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Mileage").Select Datatable("Mileage","Third Party Vehicle")
									End If
								End IF
							End IF
						End IF
					End IF
				Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Policly_Next>>").Click
				Browser("ClaimsBrowser").Sync
End Function

Rem Function Written for   vechile location 
'Location of vehicle   With InsuredWith ClaimantStorage FacilityRepairerInsured Other Than Site  

Function  ThirdParty_Owner_Vehicle_Location()
	If Datatable("Third_LossEval_LocOfVehicle","Third Party Vehicle")="With Insured" Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Location of Vehicle").Select Datatable("Third_LossEval_LocOfVehicle","Third Party Vehicle")
	End IF
	
	If Datatable("Third_LossEval_LocOfVehicle","Third Party Vehicle")="With Claimant"  then 
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Location of Vehicle").Select Datatable("Third_LossEval_LocOfVehicle","Third Party Vehicle")	
	End IF
	
	If Datatable("Third_LossEval_LocOfVehicle","Third Party Vehicle")="Storage Facility" Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Location of Vehicle").Select Datatable("Third_LossEval_LocOfVehicle","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Name").Set Datatable("Storage Facility_Name","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Address1").Set Datatable("Storage Facility_Address1","Third Party Vehicle")	
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Address2").Set  Datatable("Storage Facility_Address2","Third Party Vehicle")	
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Zip").Set Datatable("Storage Facility_Zip","Third Party Vehicle")	
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_PPhone").Set Datatable("Storage Facility_PPhone","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_APhone").Set  Datatable("Storage Facility_APhone","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Email").Set  Datatable("Storage Facility_Email","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Fax").Set Datatable("Storage Facility_Fax","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Storage Facility_Permission to Release").Select Datatable("Storage Facility_Permission to Release","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Storage Facility_Tow bill").Select Datatable("Storage Facility_Tow bill","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Vehicle").Set Datatable("Storage Facility_Vehicle","Third Party Vehicle")	
	End If

	If Datatable("Third_LossEval_LocOfVehicle","Third Party Vehicle")="Repairer" Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information_Name").Set Datatable("Repairer Information_Name","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information_Address1").Set Datatable("Repairer Information_Address1","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information _Address2").Set Datatable("Repairer Information _Address2","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information_Zip").Set Datatable("Repairer Information_Zip","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information_PPhone").Set Datatable("Repairer Information_PPhone","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information_Email").Set Datatable("Repairer Information_Email","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Repairer Information _Fax").Set Datatable("Repairer Information _Fax","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Storage Facility_Permission to Release").Select Datatable("Storage Facility_Permission to Release","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Storage Facility_Tow bill").Select Datatable("Storage Facility_Tow bill","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Storage Facility_Vehicle").Set Datatable("SStorage Facility_Vehicle","Third Party Vehicle")
	End IF

	If Datatable("Third_LossEval_LocOfVehicle","Third Party Vehicle")="Insured Other Than Site" Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Name").Set Datatable("Insured_Other_Than_Site_Name","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Add1").Set Datatable("Insured_Other_Than_Site_Add1","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Add2").Set Datatable("Insured_Other_Than_Site_Add2","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_PPhone").Set Datatable("Insured_Other_Than_Site_Pphone","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Aphone").Set Datatable("Insured_Other_Than_Site_Aphone","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Zip").Set Datatable("Insured_Other_Than_Site_Zip","Third Party Vehicle")
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebEdit("Insured_Other_Than_Site_Email").Set  Datatable("Insured_Other_Than_Site_Email","Third Party Vehicle")
	End IF
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebButton("Policly_Next>>").Click
	
End Function 

Function Auto_ExcelReport_Generation()

	TestDataPath = Environment.Value("ClaimNumberPath") & "\Claim_Numbers\ClaimNumbers.xlsx"
	Set TestData_ExcelObj = CreateObject("Excel.Application")
	TestData_ExcelObj.Workbooks.Open (TestDataPath)
	TestData_ExcelObj.Visible=True
	Set TDSheet = TestData_ExcelObj.Sheets.Item(1)
	RowNum_i=-1
	RowNum_i = TDSheet.usedrange.rows.count  
	TDSheet.Cells(RowNum_i+1,1) =  Environment.Value("SceNum")
	TDSheet.Cells(RowNum_i+1,2) =  Environment.Value("Claim_Number_1")
	TDSheet.Cells(RowNum_i+1,3) =  Environment.Value("Claim_Number_2")
	TDSheet.Cells(RowNum_i+1,4) =  Environment.Value("SCaseId")
	TDSheet.Cells(RowNum_i+1,5) =  Date
	TDSheet.Cells(RowNum_i+1,6) =  Environment.value("str_Exe_Status")
	TDSheet.Cells(RowNum_i+1,7) =  Environment.value("str_Exe_Time")
	TestData_ExcelObj.ActiveWorkbook.Save
	TestData_ExcelObj.Application.Quit
	Set TestData_ExcelObj = Nothing
	
End Function


Function Auto_Update_Regression_Tracker()
	
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
		excelSheet.Cells(FindCell.row,3) =  Environment.Value("Claim_Number_1")
		excelSheet.Cells(FindCell.row,4) =  Environment.Value("Claim_Number_2")
		excelObj.Activeworkbook.Save
	
	ElseIf NOT FindCell2 is Nothing  Then
		excelSheet2.Cells(FindCell2.row,2) = Environment.Value("SCaseId")
		excelSheet.Cells(FindCell.row,3) =  Environment.Value("Claim_Number_1")
		excelSheet.Cells(FindCell.row,4) =  Environment.Value("Claim_Number_2")
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

