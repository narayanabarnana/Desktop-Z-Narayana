'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
								'Carepoint NS Business Functions
								'Updated By : Srirekha Talasila
								'Updated On : 12/14/2016
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


 Function Login()

	Environment.value("str_ScreenName") = "Carepoint >>>> NS - Login Page "
	SystemUtil.CloseProcessByName "iexplore.exe"
	Systemutil.Run "iexplore.exe", Environment.Value("CP_URL")	
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("UserIdentifier").Set Environment.Value("CP_LoginId")
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("Password").Set Environment.Value("CP_LoginPassword") 
	Browser("ClaimsBrowser").Page("LoginPage").WebButton("Log In").Click
	
 End function

Function Select_NS()

	Environment.value("str_ScreenName") = "Carepoint >>>> NS - Select NS "
	Browser("name:=CCC.*").Page("title:=CCC.*").Link("html tag:=A","name:=My Group").Click
	Browser("name:=CCC.*").Page("title:=CCC.*").WebList("html id:=objWorkBasketSelect","html tag:=SELECT","name:=select").Select "Non-Standard"
	
End function

Function Select_Workitem()

	Environment.value("str_ScreenName") = "Carepoint >>>> NS - Select WorkItem "
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

	Environment.value("str_ScreenName") = "Carepoint >>>> NS - Customer Search "
	Dim objBrwpage_CustomerSearch

	set objBrwpage_CustomerSearch=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame") 
	objBrwpage_CustomerSearch.WebButton("Customer Search").Click 
	Wait(3)
	If (DataTable("Search_Flow","GL-Data") = "Customer") Then
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
	ElseIf DataTable("Search_Flow","GL-Data") = "Employee" Then		
		Employee_Search()
	Else
		'Do Nothing
	End If
	
End Function


Function Add_NewCustomer()

		Environment.value("str_ScreenName") = "Carepoint - NS >>>> Add New Customer "
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Add Customer").Click
		Browser("ClaimsBrowser").sync
        Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Name").Set DataTable("AddCust_Name","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Address").Set DataTable("AddCust_Address","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Zip").Set DataTable("AddCust_Zip","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Phone").Set DataTable("AddCust_Phone","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Email").Set DataTable("AddCust_Email","GL-Data")
'		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("html id:=taxId").Set DataTable("AddCustomer_EmpTaxID","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_SiteCode").Set DataTable("CS_SiteCode","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("AddCust_Submit").Click
		wait(2)
		If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(5) then
				Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click
		Else
			If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist(5) then
				Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
			End If 
		End If
		If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Click
		End If
		
End Function


Function Employee_Search()

	Environment.value("str_ScreenName") = "Carepoint - NS >>>>  Employee Search "
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Employee Search").Click
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Emp_CustomerName").Exist(8) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Emp_CustomerName").Select DataTable("Emp_CustomerName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Emp_Search").Click
	End If
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebRadioGroup("Emp_Result").Exist(15) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebRadioGroup("Emp_Result").Select "1"
	End If
    Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Emp_Select").Click
    wait 1
	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("actionIFrame").WebButton("Start Process").Exist(5) Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("actionIFrame").WebButton("Start Process").Click
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
	
	Environment.value("str_ScreenName") = "Carepoint - NS >>>> Incident Screen "
	Set NS_Incident=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")  
	NS_Incident.WebEdit("IN_AccidentDate").Set DataTable("CS_Accident_Date","GL-Data")
	NS_Incident.WebList("AccidentState").Select DataTable("IN_AccidentState","GL-Data")
	NS_Incident.WebList("AccidentTime1").Select DataTable("IN_AccidentTime1","GL-Data")
	NS_Incident.WebList("AccidentTime2").Select DataTable("IN_AccidentTime2","GL-Data")
	NS_Incident.WebList("AccidentTime3").Select DataTable("IN_AccidentTime3","GL-Data")
	NS_Incident.WebList("Catagory").Select DataTable("IN_Category","GL-Data")
	NS_Incident.WebList("IN_Product").Select DataTable("IN_Product","GL-Data")
	If  NS_Incident.WebList("IN_Exposure").Exist(5) Then
		NS_Incident.WebList("IN_Exposure").Select DataTable("IN_Exposure","GL-Data")
	End If
	If DataTable("IN_Product","GL-Data") = "Reinsurance" or DataTable("IN_Product","GL-Data") = "SAFE" or DataTable("IN_Product","GL-Data") = "Occupational Accident"  Then 
	 	NS_Incident.WebEdit("IN_DateReported").Set DataTable("IN_DateReported","GL-Data")
	End If
	If  DataTable("IN_Product","GL-Data") = "Surety/Fidelity" Then
		NS_Incident.WebEdit("IN_DateReported").Set DataTable("IN_DateReported","GL-Data")
		NS_Incident.WebCheckBox("Is_9_Series").Set DataTable("Is_9_Series","GL-Data")
	End If
	If  DataTable("IN_Product","GL-Data")="Jockey" Then
		NS_Incident.WebEdit("IN_DateReported").Set DataTable("IN_DateReported","GL-Data")
	End If
	NS_Incident.WebEdit("IN_Claimant_Fname").Set DataTable("IN_Fname","GL-Data")
	NS_Incident.WebEdit("IN_Claimant_MI").Set DataTable("IN_MI","GL-Data") 
	NS_Incident.WebEdit("IN_Claimant_Lname").Set DataTable("IN_Lname","GL-Data")
	NS_Incident.WebEdit("AccidentDescription").Set DataTable("IN_AccDescription","GL-Data")
	
	NS_Incident.WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	
	If  NS_Incident.WebButton("No Duplicates Found").Exist(5) Then 'If Duplicate Claim Exists
		 NS_Incident.WebButton("No Duplicates Found").Click
		 Browser("ClaimsBrowser").sync
	End If
	
	If DataTable("IN_Product","GL-Data")="Jockey" Then
		NS_Incident.Link("title:=Jockey - Policy and Address ").Click
		wait(4)
		Browser("ClaimsBrowser").WinObject("text:=Do you want to open Jockey\.xlsx \(20\.7 KB\) from teamspace\.zurichna\.com\?").WinButton("acc_name:=Open").Click
		wait(10)
'		Window("Windows Internet Explorer").Dialog("Windows Internet Explorer").WinButton("Open").Click
	End If
	
	If Browser("ClaimsBrowser").Dialog("Use_HC_PolicyOnly").Exist(5) then
		Browser("ClaimsBrowser").Dialog("Use_HC_PolicyOnly").WinButton("OK").Click
	End If
	
	
End Function
 
Function PolicySearch()

	Environment.value("str_ScreenName") = "Carepoint - NS >>>> Policy Search Screen "
	
	If DataTable("IN_Product","GL-Data") = "Occupational Accident"  Then
	   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PS_PolicyNumber").Set DataTable("PS_PolicyNo","GL-Data")
	   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Policy_Retrieve").Click
	   Wait(3)
	   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("Policy_RadioButton").Click
	   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	Else	
		If DataTable("CS_Policynum","GL-Data")<>"" Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PS_PolicyNumber").Set DataTable("CS_Policynum","GL-Data")	
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Policy_Retrieve").Click
			Browser("ClaimsBrowser").sync
		End If 	
			cell_data = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebTable("Policy_Table").GetCellData(2,1)
			If cell_data = "" Then
				Set polobj = browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebTable("Policy_Table")
				Set polobj2 = polobj.ChildItem(2,1,"WebRadioGroup",0)
				d = polobj2.getroproperty("class")
				If d = "Radio lvInputSelection" Then
					Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("Policy_RadioButton").Click
					Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
				End if
			Else
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("Policy_RadioButton").highlight
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("Policy_RadioButton").Select "Indeterminate"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
				Browser("ClaimsBrowser").sync
			End If			
	End If

	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist(10) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Click
		Browser("ClaimsBrowser").sync
	End if
	Browser("ClaimsBrowser").Sync	
End Function

Function Override_TPA()

	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Exist(6) then
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Click
	Else
		'Do Nothing
	End If
	
	

End Function

Function Contact_Info()

	Environment.value("str_ScreenName") = "Carepoint - NS >>>> Contact Info Screen "
	
	Set NS_ConInfo=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	NS_ConInfo.WebEdit("CO_Rep_Name").Set DataTable("CO_Rep_Name","GL-Data")
	NS_ConInfo.WebEdit("CO_Rep_Email").Set DataTable("CO_Rep_Email","GL-Data")
	NS_ConInfo.WebEdit("CO_Rep_Phone").Set DataTable("CO_Rep_Phone","GL-Data")
	NS_ConInfo.WebList("CO_Report_Relation").Select DataTable("CO_Rep_Relation","GL-Data")
	NS_ConInfo.WebEdit("CO_CusCon_Email").Set DataTable("CO_CusCon_Email","GL-Data")
	NS_ConInfo.WebEdit("CO_CusCon_Fax").Set DataTable("CO_CusCon_Fax","GL-Data")		
	NS_ConInfo.WebButton("Next>>").Click
	
End function

Function Accident_Page()

	Environment.value("str_ScreenName") = "Carepoint - NS >>>> Accident Screen "
	Browser("ClaimsBrowser").Sync
	Set NS_Accident=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")  
	If DataTable("IN_Product","GL-Data")="Jockey" OR DataTable("IN_Product","GL-Data")="Commercial Tank Pull" Then
	    	' Do Nothing
	Else
		NS_Accident.WebList("ACC_AccCode").Select DataTable("ACC_AccCode","GL-Data")
		NS_Accident.WebList("ACC_AgentLoss").Select DataTable("ACC_AgentLoss","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Click
		Browser("ClaimsBrowser").Sync
		wait 3
		NS_Accident.WebList("ACC_LossLoc").Select DataTable("ACC_LossLoc","GL-Data")
	End If
	
	If DataTable("IN_Product","GL-Data") = "Occupational Accident" Then 
		Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*").WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAccident.*pBenState","index:=1").Select "DBA" 
	End If
	
	If  DataTable("ACC_SiteAddress","GL-Data")="No" Then
	 	NS_Accident.WebEdit("ACC_AccAddress1").Set DataTable("ACC_AccAddress1","GL-Data")
	 	NS_Accident.WebEdit("ACC_AccAddress2").Set DataTable("ACC_AccAddress2","GL-Data")
	 	NS_Accident.WebEdit("ACC_AccZip").Set DataTable("ACC_AccZip","GL-Data")
	Else
		NS_Accident.WebList("ACC_SiteAddress").Select DataTable("ACC_SiteAddress","GL-Data")
	End If
	
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("title:=FNOL.*").WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAccident.*pAddr.*pState").GetROProperty("value")="Select..." Then ''''no value exist in the Zip code
		Browser("ClaimsBrowser").Page("Inbox").Frame("title:=FNOL.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAccident.*pAddr.*ppostalCode").Set DataTable("ACC_AccZip","GL-Data")
	End If
	
	If DataTable("IN_Product","GL-Data") = "Reinsurance" or DataTable("IN_Product","GL-Data") = "SAFE" or DataTable("IN_Product","GL-Data") = "Surety/Fidelity" Then
		'do nothing
	Else
		NS_Accident.WebEdit("ACC_Comments").Set DataTable("ACC_Comments","GL-Data")
		NS_Accident.WebCheckBox("ACC_Police").Set DataTable("ACC_Police","GL-Data")
		NS_Accident.WebCheckBox("ACC_Fire").Set DataTable("ACC_Fire","GL-Data")
		NS_Accident.WebCheckBox("ACC_Ambulance").Set DataTable("ACC_Ambulance","GL-Data")
	 	NS_Accident.WebCheckBox("ACC_Other").Set DataTable("ACC_Other","GL-Data")
		If DataTable("ACC_Police","GL-Data") = "ON" Then
			NS_Accident.WebEdit("ACC_Pol_AuthName").Set DataTable("ACC_Pol_AuthName","GL-Data")
			NS_Accident.WebEdit("ACC_Pol_OffName").Set DataTable("ACC_Pol_OffName","GL-Data")
			NS_Accident.WebEdit("ACC_Pol_OffBatch").Set DataTable("ACC_Pol_OffBatch","GL-Data")
			NS_Accident.WebEdit("ACC_Pol_Report").Set DataTable("ACC_Pol_Report","GL-Data")
			NS_Accident.WebEdit("ACC_Pol_OffPhone").Set DataTable("ACC_Pol_OffPhone","GL-Data")
			NS_Accident.WebEdit("ACC_Pol_NCIC").Set DataTable("ACC_Pol_NCIC","GL-Data")
		ElseIf ((DataTable("ACC_Fire","GL-Data") = "ON") OR (DataTable("ACC_Ambulance","GL-Data") = "ON") OR (DataTable("ACC_Other","GL-Data") = "ON")) Then
			NS_Accident.WebEdit("ACC_Ambu_AuthName").Set DataTable("ACC_Ambu_AuthName","GL-Data")
			NS_Accident.WebEdit("ACC_Ambu_Report").Set DataTable("ACC_Ambu_Report","GL-Data")
			NS_Accident.WebEdit("ACC_Ambu_OSHA").Set DataTable("ACC_Ambu_OSHA","GL-Data")
		End If
	End If
	
	NS_Accident.WebButton("Next>>").Click 
	

End function


 
Function Party()

	Environment.value("str_ScreenName") = "Carepoint - NS >>>> Party Screen "
	
	Dim objBrwpage_Party,i
	set objBrwpage_Party=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
'	DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
	If objBrwpage_Party.WebEdit("Party_Fname").Exist(10) Then
		
		i = 1
		counter = Environment.Value("counter")
		If DataTable("Different_Claimant_Data","GL-Data") = "Yes" then
			For i = 1 to DataTable("No.Of.Claimants","GL-Data")
				
				objBrwpage_Party.WebEdit("Party_Fname").Set DataTable("Party_Fname","GL-Data")
				objBrwpage_Party.WebEdit("Party_MI").Set DataTable("Party_MI","GL-Data")
				objBrwpage_Party.WebEdit("Party_Lname").Set DataTable("Party_Lname","GL-Data")
				
				If DataTable("Party_Injured","GL-Data") = "ON" Then
					objBrwpage_Party.WebCheckBox("Party_Injured").Set DataTable("Party_Injured","GL-Data")	
				End If
				
				If DataTable("Party_Fatality","GL-Data") = "ON" Then
					objBrwpage_Party.WebCheckBox("Party_Fatality").Set DataTable("Party_Fatality","GL-Data")
				End If
				
				If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
					objBrwpage_Party.WebCheckBox("Party_PropertyDamage").Set DataTable("Party_PropertyDamage","GL-Data")
				End If
				
				If DataTable("Party_Witness","GL-Data") = "ON" Then
					objBrwpage_Party.WebCheckBox("Party_Witness").Set DataTable("Party_Witness","GL-Data")
				End If
				
				objBrwpage_Party.WebButton("name:=Add To List").Click
				Browser("ClaimsBrowser").Sync
				
			Next	
		Else
			If  DataTable("No.Of.Claimants","GL-Data") = "1" Then
				If objBrwpage_Party.WebElement("Ram").Exist(5) then
					objBrwpage_Party.WebElement("Ram").Click
					objBrwpage_Party.WebButton("Delete From List").Click
					Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
				End If
			End If
			For i = 1 to DataTable("No.Of.Claimants","GL-Data")
				objBrwpage_Party.WebEdit("Party_Fname").Set DataTable("Party_Fname","GL-Data")
				objBrwpage_Party.WebEdit("Party_MI").Set DataTable("Party_MI","GL-Data")
				objBrwpage_Party.WebEdit("Party_Lname").Set DataTable("Party_Lname","GL-Data")
				If DataTable("Party_Injured","GL-Data") = "ON" Then
					objBrwpage_Party.WebCheckBox("Party_Injured").Set DataTable("Party_Injured","GL-Data")	
				End If
				If DataTable("Party_Fatality","GL-Data") = "ON" Then
					objBrwpage_Party.WebCheckBox("Party_Fatality").Set DataTable("Party_Fatality","GL-Data")
				End If
				
				If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
					objBrwpage_Party.WebCheckBox("Party_PropertyDamage").Set DataTable("Party_PropertyDamage","GL-Data")
				End If
				If DataTable("Party_Witness","GL-Data") = "ON" Then
					objBrwpage_Party.WebCheckBox("Party_Witness").Set DataTable("Party_Witness","GL-Data")
				End If
				objBrwpage_Party.WebButton("name:=Add To List").Click
				Browser("ClaimsBrowser").Sync
			Next
		End if			
	End If
	
	objBrwpage_Party.WebButton("Next>>").Click
	
End function


Function Employment()  ''This function newly added after  Occupational Accident 

	Environment.value("str_ScreenName") = "Carepoint - NS >>>> Employment Screen "
	
	Set Obj_Employment=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployerAddress.*ppostalCode").Set "12345"
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pOccupation").Set DataTable("Employee_RegularOccupation","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pdeptNumber").Set DataTable("Employee_Dept","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pStartDate").Set DataTable("Employee_HireDate","GL-Data")
	Obj_Employment.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pemploymentStatus").Select DataTable("Employee_Status","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pSupervisorName").Set DataTable("Employee_SupervisorName","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pTelNbr.*gPhone.*pPhone").Set DataTable("Employee_SupervisorPhone","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pemployerNotifiedDate").Set DataTable("Employee_NotifiedDate","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pwageAmt").Set DataTable("Employee_WageAmount","GL-Data")
	Obj_Employment.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*psalaryPaymentFrequency").Select DataTable("Employee_Hourly","GL-Data")
	Obj_Employment.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pworkHoursPerDay").Set DataTable("Employee_Hours","GL-Data")
	Obj_Employment.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pworkHoursPerWeek").Select DataTable("Employee_Days","GL-Data")
	Obj_Employment.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pWorkShift").Select DataTable("Employee_WorkShift","GL-Data")
	Obj_Employment.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*plostTimeIndicator").Select DataTable("Employee_LostTime","GL-Data")
	
End Function

													
Function PartyInfo1()

		Environment.value("str_ScreenName") = "Carepoint - NS >>>> First Party Screen "
		counter = Environment.Value("counter")
		DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
		If DataTable("Different_Claimant_Data","GL-Data") = "Yes" then
			For i = 1 to DataTable("No.Of.Claimants","GL-Data")
					Environment.value("str_ScreenName") = "Carepoint - NS >>>> First Party Screen "
					If i=1 Then
						If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Exist(5) Then
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
						End If
					Else
					    Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No  
                     End If
                     
					 If DataTable("IN_Product","GL-Data")  = "Reinsurance" or DataTable("IN_Product","GL-Data")  = "HBP-HVP" Then
								' DoNothing
					 else
						If i=1 then
								call Attorney()
							Else	'''here i=2/3 i.e in Attorney page objects is unable to identify dude to the over lap of objects in i=1  
'								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("Party2_AttorneyList").Select DataTable("Attorney_List","GL-Data")
'								If DataTable("Attorney_List","GL-Data") = "Yes" Then
'									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Party2_Att_FirmName").Set DataTable("Attorney_FirmName","GL-Data")
'									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Party2_Att_FirstName").Set DataTable("Attorney_FirstName","GL-Data")
'									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Party2_Att_LastName").Set DataTable("Attorney_LastName","GL-Data")
'									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Party2_Att_Address1").Set DataTable("Attorney_Address1","GL-Data")
'									Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebEdit("Party2_Att_ZIP").Set DataTable("Attorney_ZIP","GL-Data")
'								 End If	
								call Attorney()
								If DataTable("IN_Product","GL-Data") = "Occupational Accident" And Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pEmployeeId").Exist(5) Then
									Call Employment()
								End If 
							 End If 
						End If
						If DataTable("IN_Product","GL-Data") = "Occupational Accident" And Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pEmployer.*pEmployment.*pEmployeeId").Exist(5) Then
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
							End if 
						
'''''''''''''''''''''''''''''''''''''''''''''''''''Injury  Info1''''''''''''''''''''''''''''''''''''''''''''''''''
							Environment.value("str_ScreenName") = "Carepoint - NS >>>> First Party Injury Screen "
							If i=1 Then
								If DataTable("Party_Injured","GL-Data") = "ON" Then
									ForFirstClaimntInj=DataTable("Party_Injured","GL-Data")
									Browser("ClaimsBrowser").Sync
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description1").Set DataTable("Inj_Description","GL-Data")
									If DataTable("Party_Fatality","GL-Data") = "ON" Then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_DateOfDeath").Set DataTable("Inj_DateOfDeath","GL-Data")
									End If
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_CauseInjury1").Set DataTable("Inj_CauseInjury","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_Nature1").Select DataTable("Inj_Nature","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_BodyPart1").Select  DataTable("Inj_BodyPart","GL-Data")
									If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment1").Exist(5) then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment1").Select DataTable("Inj_InitialTreatment","GL-Data")
									End if 
									If DataTable("IN_Product","GL-Data") = "Jockey" Then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=Occupation").Set "Software Engineer"
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=InjOccupation").Set "Injury Occu"
									End If
									If DataTable("Inj_InitialTreatment","GL-Data") = "NO MEDICAL TREATMENT"  or DataTable("Inj_InitialTreatment","GL-Data") = "MINOR ON-SITE REMEDIES BY EMPLOYER MEDICAL STAFF" Then
							            'do nothing
									else
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_LName").Set DataTable("Inj_Phy_LName", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_FName").Set DataTable("Inj_Phy_FName", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_MI").Set DataTable("Inj_Phy_MI", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Add1").Set DataTable("Inj_Phy_Add1", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Add2").Set DataTable("Inj_Phy_Add2", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Zip").Set DataTable("Inj_Phy_Zip", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Phone").Set DataTable("Inj_Phy_Phone", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Fax").Set DataTable("Inj_Phy_Fax", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Email").Set DataTable("Inj_Phy_Email", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Name").Set DataTable("Inj_Hosp_Name", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Add1").Set DataTable("Inj_Hosp_Add1", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Add2").Set DataTable("Inj_Hosp_Add2", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Zip").Set DataTable("Inj_Hosp_Zip", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Phone").Set DataTable("Inj_Hosp_Phone", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Fax").Set DataTable("Inj_Hosp_Fax", "GL-Data")
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Email").Set DataTable("Inj_Hosp_Email", "GL-Data")
									End if	
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
								End If
						Else
						
							DataTable("Party_Injured","GL-Data")=ForFirstClaimntInj
							
							If DataTable("Party_Injured","GL-Data") = "ON" And Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description1").Exist(5) Then							
									Environment.value("str_ScreenName") = "Carepoint - NS >>>> First Party Injury Screen "
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description1").Set DataTable("Inj_Description","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_CauseInjury1").Set DataTable("Inj_CauseInjury","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_Nature1").Select DataTable("Inj_Nature","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_BodyPart1").Select  DataTable("Inj_BodyPart","GL-Data")
									If DataTable("Party_Fatality","GL-Data") = "ON" Then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_DateOfDeath").Set DataTable("Inj_DateOfDeath","GL-Data")
									End If
									If DataTable("IN_Product","GL-Data") = "Jockey" Then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=Occupation").Set "Software Engineer"
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=InjOccupation").Set "Injury Occu"
									End If
									If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment1").Exist(5) then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment1").Select DataTable("Inj_InitialTreatment","GL-Data")
									End if 
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
							End If 
						End If 
						
						Environment.value("str_ScreenName") = "Carepoint - NS >>>> First Party Property Damage Screen "
						
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Property Damage1'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						If i=1 Then
							  If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
							  		ForFirstClaimntDmg=DataTable("Party_PropertyDamage","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PropertyDam1_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropertyDam1_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
									If DataTable("IN_Product","GL-Data") = "Surety/Fidelity" Then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PD_ClaimType").Select DataTable("PropertyDam_ClaimType","GL-Data")
									End if 
									If DataTable("PropertyDam_ClaimType","GL-Data") = "Surety" or DataTable("PropertyDam_ClaimType","GL-Data") = "Fidelity" Then
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Name").Set "test"
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add1").Set "123main st"
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add2").Set "456main st"
										Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Zip").Set "12345"
                                        Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Phone").Set "111-222-3333"
									End if 
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
								End If
						Else '''i=2/3   '''In the OR the property Dameage objects in pro1 and prop 2 are overlapped 							
						     DataTable("Party_PropertyDamage","GL-Data")=ForFirstClaimntDmg
						     
						     If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PropertyDam1_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropertyDam1_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
								
								If DataTable("IN_Product","GL-Data") = "Surety/Fidelity" Then
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PD_ClaimType").Select DataTable("PropertyDam_ClaimType","GL-Data")
								End if 
								
								If DataTable("PropertyDam_ClaimType","GL-Data") = "Surety" or DataTable("PropertyDam_ClaimType","GL-Data") = "Fidelity" Then
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Name").Set "test"
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add1").Set "123main st"
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add2").Set "456main st"
									Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Zip").Set "12345"
                                    Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Phone").Set "111-222-3333"
								End if 
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
'								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebCheckBox("ProDam2_PropertyDam1_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
'								Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebCheckBox("ProDam2_PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
'								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
							End If  
						End if 							
									
						If DataTable("Party_Witness","GL-Data") = "ON" Then
							Call Witness()
						End If
					Next
			
		Else
					Environment.value("str_ScreenName") = "Carepoint - NS >>>> First Party Screen "
					If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Exist(5) Then
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
					End If
					If DataTable("IN_Product","GL-Data")  <> "Reinsurance" or DataTable("IN_Product","GL-Data")  = "HBP-HVP" Then
						call Attorney()
					End If						
					Browser("ClaimsBrowser").Sync
					If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo1_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
					End If		
					Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
					
					Environment.value("str_ScreenName") = "Carepoint - NS >>>> First Party Injury Screen "
					''Injury  Info1'
					If DataTable("Party_Injured","GL-Data") = "ON" Then
						Browser("ClaimsBrowser").Sync
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description1").Set DataTable("Inj_Description","GL-Data")
						If DataTable("Party_Fatality","GL-Data") = "ON" Then
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_DateOfDeath").Set DataTable("Inj_DateOfDeath","GL-Data")
						End If				
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_CauseInjury1").Set DataTable("Inj_CauseInjury","GL-Data")
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description1").Set DataTable("Inj_Description","GL-Data")
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_Nature1").Select DataTable("Inj_Nature","GL-Data")
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_BodyPart1").Select  DataTable("Inj_BodyPart","GL-Data")
						If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment1").Exist(5) then
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment1").Select DataTable("Inj_InitialTreatment","GL-Data")
						End if 
									
'						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("Inj_Nature1").Select DataTable("Inj_Nature","GL-Data")
'						Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("Inj_BodyPart1").Select DataTable("Inj_BodyPart","GL-Data")
'						If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("Inj_InitialTreatment1").Exist(5) then
'							Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("PartyInfo2").WebList("Inj_InitialTreatment1").Select DataTable("Inj_InitialTreatment","GL-Data")
'						End if 
						If DataTable("Inj_InitialTreatment","GL-Data") = "NO MEDICAL TREATMENT"  or DataTable("Inj_InitialTreatment","GL-Data") = "MINOR ON-SITE REMEDIES BY EMPLOYER MEDICAL STAFF" Then
							'do nothing
						else						
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_LName").Set DataTable("Inj_Phy_LName", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_FName").Set DataTable("Inj_Phy_FName", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_MI").Set DataTable("Inj_Phy_MI", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Add1").Set DataTable("Inj_Phy_Add1", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Add2").Set DataTable("Inj_Phy_Add2", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Zip").Set DataTable("Inj_Phy_Zip", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Phone").Set DataTable("Inj_Phy_Phone", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Fax").Set DataTable("Inj_Phy_Fax", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Phy_Email").Set DataTable("Inj_Phy_Email", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Name").Set DataTable("Inj_Hosp_Name", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Add1").Set DataTable("Inj_Hosp_Add1", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Add2").Set DataTable("Inj_Hosp_Add2", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Zip").Set DataTable("Inj_Hosp_Zip", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Phone").Set DataTable("Inj_Hosp_Phone", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Fax").Set DataTable("Inj_Hosp_Fax", "GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Hosp_Email").Set DataTable("Inj_Hosp_Email", "GL-Data")							
						End if	
						Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
					End If

					Environment.value("str_ScreenName") = "Carepoint - NS >>>> First Party PD Screen "
					'Property Damage1
						Browser("ClaimsBrowser").Sync
						If DataTable("Party_PropertyDamage","GL-Data") = "ON" and Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_PropDescription").Exist(5)  Then
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PropertyDam1_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam1_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropertyDam1_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
							If DataTable("IN_Product","GL-Data") = "Surety/Fidelity" Then
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PD_ClaimType").Select DataTable("PropertyDam_ClaimType","GL-Data")
							End if 
							If DataTable("PropertyDam_ClaimType","GL-Data") = "Surety" or DataTable("PropertyDam_ClaimType","GL-Data") = "Fidelity" Then
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Name").Set "test"
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add1").Set "123main st"
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Add2").Set "456main st"
								Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Zip").Set "12345"
                                Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PD_ClaimType_Phone").Set "111-222-3333"
							End if 
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
							Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
						End If
    			End If
    		
End Function

Function PartyInfo2()
		
		Environment.value("str_ScreenName") = "Carepoint - NS >>>> Second Party Screen "
		'Party Info 2
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo2_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
		If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo2_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo2_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo2_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
		End if
		
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Environment.value("str_ScreenName") = "Carepoint - NS >>>> Second Party Injury Screen "
		'Injury 2
		If DataTable("Party_Injured","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description2").Set DataTable("Inj_Description","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_CauseInjury2").Set DataTable("Inj_CauseInjury","GL-Data")
			If DataTable("Party_Fatality","GL-Data") = "ON" Then
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_DateOfDeath").Set DataTable("Inj_DateOfDeath","GL-Data")
			End If
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_Nature2").Select DataTable("Inj_Nature","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_BodyPart2").Select  DataTable("Inj_BodyPart","GL-Data")
			If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment2").Exist(5) then
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment2").Select DataTable("Inj_InitialTreatment","GL-Data")
			End if 
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		End  if
		 Environment.value("str_ScreenName") = "Carepoint - NS >>>> Secoond Party PD Screen "
		'Property Damage 2
		If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PropertyDam2_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam2_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam2_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam2_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropertyDam2_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam2_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam2_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		End if

End Function

Function PartyInfo3()

		Environment.value("str_ScreenName") = "Carepoint - NS >>>> Third Party Screen "
		'Party Info 3
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo3_ZurichEmp").Select DataTable("PartyInfo_ZurichEmp","GL-Data")  'value  true/false
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PartyInfo3_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
		
		If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo3_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo3_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PartyInfo3_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
		End if
		Browser("ClaimsBrowser").Sync
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		
		Environment.value("str_ScreenName") = "Carepoint - NS >>>> Third Party Injury Screen "
		'Injury 3
		If DataTable("Party_Injured","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_Description3").Set DataTable("Inj_Description","GL-Data")
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_CauseInjury3").Set DataTable("Inj_CauseInjury","GL-Data")
				If DataTable("Party_Fatality","GL-Data") = "ON" Then
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_DateOfDeath").Set DataTable("Inj_DateOfDeath","GL-Data")
				End If
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_Nature3").Select DataTable("Inj_Nature","GL-Data")
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_BodyPart3").Select  DataTable("Inj_BodyPart","GL-Data")
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Inj_InitialTreatment3").Select DataTable("Inj_InitialTreatment","GL-Data")
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		End if
		
		Environment.value("str_ScreenName") = "Carepoint - NS >>>> Third Party PD Screen "
		'Property Damage 3
		If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
				Browser("ClaimsBrowser").Sync
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PropertyDam3_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam3_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam3_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropertyDam3_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropertyDam3_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam3_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PropertyDam3_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		End if 
		
End function 


Function Witness()
	
	Environment.value("str_ScreenName") = "Carepoint - NS >>>> Witness Screen "
	If DataTable("Party_Witness","GL-Data") = "ON" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Wit_FirstName").Set DataTable("Witness_FirstName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Wit_LastName").Set DataTable("Witness_LastName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Wit_Address1").Set DataTable("Witness_Address1","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Wit_Zip").Set DataTable("Witness_Zip","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=HomePhone","html tag:=INPUT").Set DataTable("Witness_PrimaryPhone","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=Fax","html tag:=INPUT").Set DataTable("Witness_Fax","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Browser("ClaimsBrowser").Sync
	End If

End Function

Function Attorney()
	
	Environment.value("str_ScreenName") = "Carepoint - NS >>>> Attorney Screen "
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("AttorneyList").Select DataTable("Attorney_List","GL-Data")
	If DataTable("Attorney_List","GL-Data") = "Yes" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_FirmName").Set DataTable("Attorney_FirmName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_FirstName").Set DataTable("Attorney_FirstName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_LastName").Set DataTable("Attorney_LastName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_Address1").Set DataTable("Attorney_Address1","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_ZIP").Set DataTable("Attorney_ZIP","GL-Data")
	End If
	
	 
End Function


Function Additional_Information()

	Environment.value("str_ScreenName") = "Carepoint - NS >>>> Additional Info Screen "
	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("title:=FNOL.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pnoticeOfOccurence").Exist(5) Then
			Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("title:=FNOL.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pnoticeOfOccurence").Set DataTable("AddInfo_NoticeOfOccurance","GL-Data")
	End IF	
			
	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("title:=FNOL.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pnoticeOfClaim").Exist(5) Then 
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("title:=FNOL.*").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAdditionalInformation.*pnoticeOfClaim").Set DataTable("AddInfo_NoticeOfClaim","GL-Data")
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
		

End Function

Function Logout()
	
	Environment.value("str_ScreenName") = "Carepoint - NS  >>>> Logoff Screen "
	
	Browser("name:=CC.*").Page("title:=CC.*").Image("name:=Image","image type:=Image Link","Index:=0").Click
	Browser("name:=CC.*").Page("title:=CC.*").WebElement("innertext:=Log off","html id:=ItemMiddle").Click
	SystemUtil.CloseProcessByName "iexplore.exe"
	
End Function

Function Assignment()
	
	Environment.value("str_ScreenName") = "Carepoint - NS >>>> Assignment Screen "
	Browser("ClaimsBrowser").Sync
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Run Assignment").Exist(6) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Accident_Code").Select "#01"
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Run Assignment").Click
	End If	
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Click
	Browser("ClaimsBrowser").Sync
	
	If 	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Get_Claim_Number").Exist(30) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Get_Claim_Number").Click
		Browser("ClaimsBrowser").Sync
	End If
	
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist(10) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Click
	End If
	
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
	
	Environment.value("str_ScreenName") = "Carepoint - NS  >>>> Review Distribution Screen "
	
		On Error Resume Next

		Call GetClaimNumber()
		Browser("name:=CCC.*").Page("title:=CCC.*").Sync
		If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			''Log Off	
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
		
	
 End Function
 
 


Function Binocular_search()
	
	Dim incidentsearch,EXP_IncidentID
	Browser("ClaimsBrowser").Page("Inbox").WebElement("BinocularSearch").Click
	Browser("QA: Zurich Intranet Login").Page("SearchIncident").WebEdit("BinocularSearch_ClaimNbr").Set Environment.Value("ClaimNumber")
	Browser("QA: Zurich Intranet Login").Page("SearchIncident").WebButton("Search").Click	
	Set WshShell = CreateObject("WScript.Shell") 
	WshShell.SendKeys "%{ }" 
	WshShell.SendKeys " x" 
	Set WshShell=Nothing 	
	Browser("QA: Zurich Intranet Login").Page("SearchIncident").WebElement("Incident ID").Click
	Browser("QA: Zurich Intranet Login").Close	
	If Trim(Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("Resolved-Completed").GetROProperty("innertext"))="Resolved-Completed" Then 
		'''Do Nothing 
	Else
		EXPCase_Number=Trim(Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Link("EXP-Case").GetROProperty("text"))
	End If 
	Browser("ClaimsBrowser").Page("Inbox").WebElement("Inbox").Click
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("Queue").Select "Exception Handling"
	Browser("ACT II").Page("get worklist for selected").WebElement("SortDate").Click
	
	Set objref=createobject("Mercury.DeviceReplay")
    x=Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").ChildItem(2,3,"WebElement",0).GetRoProperty("abs_x")
    y=Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").ChildItem(2,3,"WebElement",0).GetRoProperty("abs_y")
    objref.MouseDblClick x,y,0   
    Set objref=nothing   
	
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Customer Search").Exist(5) Then		 
		Customer_Search()
	End If 
	
End Function


' ********************************** HealthCare Test Cases **********************************************************************
Function Re_select_Customer()

	 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Re-select Customer").Click
	 Wait 5

	 If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("CS_Search").Exist(5) then
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
		If  Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("VI_Reason").Exist(5) Then
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

	If Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebButton("IS_Save").Exist(5) Then
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

	If Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebList("none").Exist(5) Then
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

	If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist(5) then
		Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("Cancel").Click
		Wait 3
		ReportResult_Event micPass, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "PopUp is present when Close button is clicked"
	Else
		ReportResult_Event micFail, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "PopUp is not present when Close button is clicked"
	End if

	If Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebList("none").Exist(5) Then
		ReportResult_Event micPass, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "Page is navigated to Inbox after Cancel button is clicked from the Popup"
		Status = trim(Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebTable("Urgency").GetCellData(2,2))
		ScaseID = trim(Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebTable("Urgency").GetCellData(2,7))
		
			If Status = "Pending" and ScaseID = Environment.Value("SCaseId") Then
				Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebElement("IB_IncidentID").Click
				ReportResult_Event micPass, "Invoking Business component: TC25_Close_and_Reselect_Customer_Property_Damage1" , "WorkItem with the Status = 'Pending' and IncidentID ="& Environment.Value("SCaseId") & "is present in the Inbox page"
					If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Re-select Customer").Exist(5) Then
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

	Claim_Series = Left(Environment.Value("NewClaimNumber"),1)

	If Claim_Series = "9" Then 
			Print "Invoking Business component: Assignment , System generates Claim Series = 9 Subpath = 9"
	Else
			Print "Invoking Business component: Assignment , System did not generates Claim Series = 9 Subpath = 9"
	End If

End Function
'**********************************************************************************************************************************************************************
Function Val_ClaimSeries_6()

	Claim_Series = Left(Environment.Value("NewClaimNumber"),1)

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

	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("Policy_Override").Exist(5) then
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

		If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Confirm").Exist(5) then
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
