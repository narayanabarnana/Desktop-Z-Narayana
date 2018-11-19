'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
								'Carepoint HC Business Functions
								'Updated By : Srirekha Talasila
								'Updated On : 12/13/2016
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Function Login()

	Environment.value("str_ScreenName") = "Carepoint >>>> Login Page "
	Systemutil.CloseProcessByName "iexplore.exe"
	Systemutil.Run "iexplore.exe", Environment.Value("CP_URL")	
	Browser("ClaimsBrowser").Page("LoginPage").Sync
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("UserIdentifier").Set Environment.Value("CP_LoginId")
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("Password").Set Environment.Value("CP_LoginPassword") 
	Browser("ClaimsBrowser").Page("LoginPage").WebButton("Log In").Click
	
 End function


Function Select_HC()
	
	Environment.value("str_ScreenName") = "Carepoint >>>> HC -Select HC "
	Browser("name:=CCC.*").Page("title:=CCC.*").Link("html tag:=A","name:=My Group").Click
	Browser("name:=CCC.*").Page("title:=CCC.*").WebList("html id:=objWorkBasketSelect","html tag:=SELECT","name:=select").Select "Healthcare"
	
End function


Function Select_Workitem()

	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Select WorkItem "
'	Browser("Customer_Browser").Page("WorkList_Basket").WebElement("SortDate").Click
	wait(3)
	Browser("Customer_Browser").Page("WorkList_Basket").WebElement("title:=Click.*","Index:=12").click
	Wait(6)
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

	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Customer Search "
	Dim objBrwpage_CustomerSearch

	set objBrwpage_CustomerSearch=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame") 
	objBrwpage_CustomerSearch.WebButton("Customer Search").Click 
	Wait(3)
	If (DataTable("Add_NewCustomer_Flag","GL-Data") = "FALSE") Then
		If objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Exist(10) Then
		   objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Set CDATE(DataTable("CS_Accident_Date","GL-Data"))	
		End If		
		objBrwpage_CustomerSearch.WebEdit("CS_CustomerName").Set DataTable("CS_Customer_Name","GL-Data")
		objBrwpage_CustomerSearch.WebEdit("CS_SiteCode").Set DataTable("CS_SiteCode","GL-Data")
		objBrwpage_CustomerSearch.WebButton("CS_Search").Click		
		Wait(5)
		Index=1
		while index<>0'''Here the condition will waits till Web Table load
			If (objBrwpage_CustomerSearch.webelement("CS_No_Matching_Data").Exist ) Then
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
		
		
		If Browser("title:=Care.*").Exist(5) Then
		   Browser("title:=Care.*").Close 
   		End If  	
   		
		If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Click
		End If
		If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(5) then
			Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click
		Else
			If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist(5) then
				Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
			End If 
		End If
	ElseIf DataTable("Add_NewCustomer_Flag","GL-Data") = "TRUE" Then		
		Add_NewCustomer()
	Else
		'Do Nothing
	End If
	
End Function


Function Add_NewCustomer()

		Environment.value("str_ScreenName") = "Carepoint - HC >>>> Add New Customer "
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Add_New_Customer").Click
        Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("Addcustomer_CustomerName").Set DataTable("AddCustomer_CustomerName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCustomer_Addr1").Set DataTable("AddCustomer_Addr1","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCustomer_ZIP").Set DataTable("AddCustomer_ZIP","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCustomer_Phone").Set DataTable("AddCustomer_Phone","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCustomer_Email").Set DataTable("AddCustomer_Email","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("html id:=taxId").Set DataTable("AddCustomer_EmpTaxID","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCustomer_SiteCode").Set DataTable("AddCustomer_SiteCode","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("AddCustomer_Submit").Click
		wait(2)
		If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(5) then
				Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click
		Else
			If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist(5) then
				Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
			End If 
		End If
'		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("html id:=startProcessButton","title:=Complete this assignment","name:=.*Start Process.*").Click
		If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("name:=Start Process.*").Click
		End If
		
End Function

Function Employee_Search()

	Environment.value("str_ScreenName") = "Carepoint - HC >>>> Employee Search "
	Dim objBrwpage_Employee_Search
	set objBrwpage_Employee_Search=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame")
	If objBrwpage_Employee_Search.WebButton("Employee Search").Exist Then
		objBrwpage_Employee_Search.WebButton("Employee Search").Click
	End If
	If  objBrwpage_Employee_Search.WebList("Emp_CustomerName").Exist Then
		objBrwpage_Employee_Search.WebList("Emp_CustomerName").Select DataTable("Emp_CustomerName","GL-Data")
		objBrwpage_Employee_Search.WebButton("Emp_Search").Click
	End If
	If objBrwpage_Employee_Search.WebRadioGroup("Emp_Result").Exist Then
		objBrwpage_Employee_Search.WebRadioGroup("Emp_Result").Select "1"
	End If
	objBrwpage_Employee_Search.WebButton("Emp_Select").Click
	Wait(2)
	
	If Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("actionIFrame").WebButton("Start Process").Exist(5) Then
		Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("actionIFrame").WebButton("Start Process").Click
	End If
	If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(5) then
		Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click
	Else
		If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist(5) then
			Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
		End If 
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
	
	Dim objBrwpage_Incident
	Environment.value("str_ScreenName") = "Carepoint - HC >>>> Incident Screen "
	Set objBrwpage_Incident=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	If Browser("title:=Care.*").Exist(5) Then
	   Browser("title:=Care.*").Close 
	End If
    If len(Trim(objBrwpage_Incident.WebEdit("Site_TIN").GetROProperty("value")))<9 Then 
	   objBrwpage_Incident.WebEdit("Site_TIN").Set ""
    End If
    If len(Trim(objBrwpage_Incident.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pCustSiteLocation.*ppostalCode").GetROProperty("value")))<=0 Then 
	   objBrwpage_Incident.WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pCustSiteLocation.*ppostalCode").Set "12345"
    End If
	objBrwpage_Incident.WebEdit("IN_AccidentDate").Set DataTable("CS_Accident_Date","GL-Data")
	objBrwpage_Incident.WebEdit("IN_Date_Reporter").Set DataTable("IN_Date_Reporter","GL-Data")
	objBrwpage_Incident.WebList("AccidentState").Select DataTable("IN_AccidentState","GL-Data")
	objBrwpage_Incident.WebList("AccidentTime1").Select DataTable("IN_AccidentTime1","GL-Data")
	objBrwpage_Incident.WebList("AccidentTime2").Select DataTable("IN_AccidentTime2","GL-Data")
	objBrwpage_Incident.WebList("AccidentTime3").Select DataTable("IN_AccidentTime3","GL-Data")
	objBrwpage_Incident.WebEdit("AccidentDescription").Set DataTable("IN_AccDescription","GL-Data")
	objBrwpage_Incident.WebButton("Next>>").Click
	If objBrwpage_Incident.WebButton("No Duplicates Found").Exist(5) Then	'''If Duplicate Claim Exists			
		objBrwpage_Incident.WebButton("No Duplicates Found").Click
	Else 
		If objBrwpage_Incident.WebButton("No Duplicates Found").Exist(5) Then				
			objBrwpage_Incident.WebButton("No Duplicates Found").Click
		End If 
	End If
  	If Browser("ClaimsBrowser").Dialog("Use_HC_PolicyOnly").Exist(5) then
		Browser("ClaimsBrowser").Dialog("Use_HC_PolicyOnly").WinButton("OK").Click
	End If
	
	
End Function


Function PolicySearch()

	Environment.value("str_ScreenName") = "Carepoint - HC >>>> Policy Search "
	Dim objBrwpage_PolicySearch
	set objBrwpage_PolicySearch=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	Wait(2)
	If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(5) Then
	   Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click		   	   	    
	End If 
	
	If DataTable("CS_Policynum","GL-Data")="" Then
		cell_data = objBrwpage_PolicySearch.WebTable("Policy_Table").GetCellData(2,1)
		If cell_data = "" Then
			Set polobj = objBrwpage_PolicySearch.WebTable("Policy_Table")
			Set polobj2 = polobj.ChildItem(2,1,"WebRadioGroup",0)				
			d = polobj2.getroproperty("class")	
			If d = "Radio lvInputSelection" Then
				Wait(2)
				objBrwpage_PolicySearch.WebRadioGroup("Policy_RadioButton").Click
				objBrwpage_PolicySearch.WebButton("Next>>").Click
			Else
			'Report to html "No Policy Found
			End if
		Else
			If objBrwpage_PolicySearch.WebElement("innertext:=No matching policy records found.*","innerhtml:=No matching policy records found.*").Exist(6) Then
				objBrwpage_PolicySearch.WebRadioGroup("Indeterminate").Select "Indeterminate"
				objBrwpage_PolicySearch.WebButton("Next>>").Click			
			End If
		End if	
	Else 		
		objBrwpage_PolicySearch.WebEdit("PS_Policynum").Set DataTable("CS_Policynum","GL-Data")
		objBrwpage_PolicySearch.WebButton("Policy_Retrieve").Click
		cell_data = objBrwpage_PolicySearch.WebTable("Policy_Table").GetCellData(2,1)
		If objBrwpage_PolicySearch.WebRadioGroup("Policy_RadioButton").Exist(10)  Then
			objBrwpage_PolicySearch.WebRadioGroup("Policy_RadioButton").Click
			objBrwpage_PolicySearch.WebButton("Next>>").Click
		Else 
			objBrwpage_PolicySearch.WebRadioGroup("Indeterminate").Select "Indeterminate"
			objBrwpage_PolicySearch.WebButton("Next>>").Click
		End if
	End If
	
End Function


Function Override_TPA()
	
	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Override TPA "

   If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Exist(5) then
		 Wait(2)
		 Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Override_TPA").Click
	Else
		'Do Nothing
	End If

End Function


Function Contact_Info()
	
	Environment.value("str_ScreenName") = "Carepoint - HC >>>> Contact Info "
	Dim objBrwpage_Contact_Info
	set objBrwpage_Contact_Info=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	objBrwpage_Contact_Info.WebEdit("CO_Rep_Name").Set DataTable("CO_Rep_Name","GL-Data")
	objBrwpage_Contact_Info.WebEdit("CO_Rep_Email").Set DataTable("CO_Rep_Email","GL-Data")
	objBrwpage_Contact_Info.WebEdit("CO_Rep_Phone").Set DataTable("CO_Rep_Phone","GL-Data")
	objBrwpage_Contact_Info.WebList("CO_Report_Relation").Select DataTable("CO_Rep_Relation","GL-Data")
	objBrwpage_Contact_Info.WebEdit("CO_CusCon_Email").Set DataTable("CO_CusCon_Email","GL-Data")
	objBrwpage_Contact_Info.WebEdit("CO_CusCon_Fax").Set DataTable("CO_CusCon_Fax","GL-Data")		
	objBrwpage_Contact_Info.WebButton("Next>>").Click
	
End function

Function Accident_Page()
	
	Environment.value("str_ScreenName") = "Carepoint - HC >>>> Accident Screen "
	Dim objBrwpage_Accident_Page
	set objBrwpage_Accident_Page=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Exist(5) Then
	   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Click
    End If
	If objBrwpage_Accident_Page.WebList("ACC_AccCode").Exist(10) Then
		objBrwpage_Accident_Page.WebList("ACC_AccCode").Select DataTable("ACC_AccCode","GL-Data")
		objBrwpage_Accident_Page.WebList("ACC_AgentLoss").Select DataTable("ACC_AgentLoss","GL-Data")
		objBrwpage_Accident_Page.WebList("ACC_LossLoc").Select DataTable("ACC_LossLoc","GL-Data")
	End If	
	objBrwpage_Accident_Page.WebList("ACC_SiteAddress").Select DataTable("ACC_SiteAddress","GL-Data")
	Val=objBrwpage_Accident_Page.WebList("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pAccident.*pAddr.*pState").GetROProperty("value")
	If  Val="Select..."   Then
		objBrwpage_Accident_Page.WebEdit("ACC_AccZip").Set DataTable("ACC_AccZip","GL-Data")
	End If
	Accident_SiteAddr = DataTable("ACC_SiteAddress","GL-Data")
	If  ( Accident_SiteAddr = "No") Then
		objBrwpage_Accident_Page.WebEdit("ACC_AccAddress1").Set DataTable("ACC_AccAddress1","GL-Data")
		objBrwpage_Accident_Page.WebEdit("ACC_AccAddress2").Set DataTable("ACC_AccAddress2","GL-Data")
		Wait(2)
		objBrwpage_Accident_Page.WebEdit("ACC_AccZip").Set DataTable("ACC_AccZip","GL-Data")
	Else
		'Do Nothing
	End If
	objBrwpage_Accident_Page.WebCheckBox("ACC_ProductDefect").Set DataTable("ACC_ProductDefect","GL-Data")
	Wait 1
	If  DataTable("ACC_ProductDefect","GL-Data") = "ON" Then
		objBrwpage_Accident_Page.WebList("ACC_ProductDefect_MD").Select DataTable("ACC_ProductDefect_MD","GL-Data")
		If DataTable("ACC_ProductDefect_MD","GL-Data") = "No" Then
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_Name").Set DataTable("ACC_ProductDefect_MD_Name","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_Add1").Set DataTable("ACC_ProductDefect_MD_Add1","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_Add2").Set DataTable("ACC_ProductDefect_MD_Add2","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_Zip").Set DataTable("ACC_ProductDefect_MD_Zip","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_AltPhone").Set DataTable("ACC_ProductDefect_MD_AltPhone","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_Phone").Set DataTable("ACC_ProductDefect_MD_Phone","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_Email").Set DataTable("ACC_ProductDefect_MD_Email","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_Fax").Set DataTable("ACC_ProductDefect_MD_Fax","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_PD_ComponetDesp").Set DataTable("ACC_ProductDefect_MD_PD_ComponetDesp","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_PD_InsuCustName").Set DataTable("ACC_ProductDefect_MD_PD_InsuCustName","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_PD_Location").Set DataTable("ACC_ProductDefect_MD_PD_Location","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_PD_ProductCode").Set DataTable("ACC_ProductDefect_MD_PD_ProductCode","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_PD_ProductDesp").Set DataTable("ACC_ProductDefect_MD_PD_ProductDesp","GL-Data")
			objBrwpage_Accident_Page.WebEdit("ACC_ProductDefect_MD_PD_Serial").Set DataTable("ACC_ProductDefect_MD_PD_Serial","GL-Data")
		End If 		
	End If
	objBrwpage_Accident_Page.WebButton("Next>>").Click 

End function



Function Party()
	
	Environment.value("str_ScreenName") = "Carepoint - HC >>>> Party Screen "
	Dim objBrwpage_Party
	set objBrwpage_Party=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	
	If objBrwpage_Party.WebEdit("Party_Fname").Exist(10) Then
		Dim i
		i = "1"
		counter = Environment.Value("counter")  
		If DataTable("Different_Claimant_Data","GL-Data") = "Yes" then
			For i = 1 to DataTable("No.Of.Claimants","GL-Data")
				objBrwpage_Party.WebEdit("Party_Fname").Set DataTable("Party_Fname","GL-Data")
				objBrwpage_Party.WebEdit("Party_MI").Set DataTable("Party_MI","GL-Data")
				objBrwpage_Party.WebEdit("Party_Lname").Set DataTable("Party_Lname","GL-Data")
				objBrwpage_Party.WebCheckBox("Party_Injured").Set DataTable("Party_Injured","GL-Data")
				objBrwpage_Party.WebCheckBox("Party_Fatality").Set DataTable("Party_Fatality","GL-Data")
				objBrwpage_Party.WebCheckBox("Party_PropertyDamage").Set DataTable("Party_PropertyDamage","GL-Data")
				objBrwpage_Party.WebCheckBox("Party_Attorney").Set DataTable("Party_Attorney","GL-Data")
				objBrwpage_Party.WebButton("name:=Add To List").Click
				Browser("ClaimsBrowser").Sync
				counter =counter + 1
				DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
			Next	
		Else
			If  DataTable("No.Of.Claimants","GL-Data") = "1" Then
				If objBrwpage_Party.WebElement("Ram").Exist then
					objBrwpage_Party.WebElement("Ram").Click
					objBrwpage_Party.WebButton("Delete From List").Click
					Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("OK").Click
				End If
			End If
			For i = 1 to DataTable("No.Of.Claimants","GL-Data")
				objBrwpage_Party.WebEdit("Party_Fname").Set DataTable("Party_Fname","GL-Data")
				objBrwpage_Party.WebEdit("Party_MI").Set DataTable("Party_MI","GL-Data")
				objBrwpage_Party.WebEdit("Party_Lname").Set DataTable("Party_Lname","GL-Data")
				objBrwpage_Party.WebCheckBox("Party_Injured").Set DataTable("Party_Injured","GL-Data")
				objBrwpage_Party.WebCheckBox("Party_Fatality").Set DataTable("Party_Fatality","GL-Data")
				objBrwpage_Party.WebCheckBox("Party_PropertyDamage").Set DataTable("Party_PropertyDamage","GL-Data")
				objBrwpage_Party.WebCheckBox("Party_Attorney").Set DataTable("Party_Attorney","GL-Data")
				objBrwpage_Party.WebButton("name:=Add To List").Click
				Browser("ClaimsBrowser").Sync
			Next
		End if			
	End If
	objBrwpage_Party.WebButton("Next>>").Click
	

End function



Function PartyInfo1()
		
	Environment.value("str_ScreenName") = "Carepoint >>>> HC - First Party - Party Info Screen "		 
	Dim objBrwpage_PartyInfo1
	set objBrwpage_PartyInfo1=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	
	counter = Environment.Value("counter")

	If DataTable("Different_Claimant_Data","GL-Data") = "Yes" then
		If  objBrwpage_PartyInfo1.WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Exist(5) Then
			objBrwpage_PartyInfo1.WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
			Wait(1)
		End If
		'If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
			objBrwpage_PartyInfo1.WebEdit("PartyInfo1_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
			objBrwpage_PartyInfo1.WebEdit("PartyInfo1_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
			objBrwpage_PartyInfo1.WebEdit("PartyInfo1_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
			objBrwpage_PartyInfo1.WebEdit("PartyInfo1_PriPhone").Set DataTable("PartyInfo_PriPhone","GL-Data")
		'End If
		objBrwpage_PartyInfo1.WebEdit("PartyInfo1_Fax").Set DataTable("PartyInfo_Fax","GL-Data")
		objBrwpage_PartyInfo1.WebEdit("PartyInfo1_Email").Set DataTable("PartyInfo_Email","GL-Data")
		objBrwpage_PartyInfo1.WebEdit("PartyInfo1_DOB").Set DataTable("PartyInfo_DOB","GL-Data")
		objBrwpage_PartyInfo1.WebList("PartyInfo1_Distribution").Select DataTable("PartyInfo_Distribution","GL-Data")
		Browser("ClaimsBrowser").Sync
		objBrwpage_PartyInfo1.WebList("PartyInfo1_Gender").Select DataTable("PartyInfo_Gender","GL-Data")
		Browser("ClaimsBrowser").Sync
		Browser("ClaimsBrowser").Sync
		objBrwpage_PartyInfo1.WebButton("Next>>").Click
		'''''''''''''''''''''''''''''''''' Injury  Info1 '''''''''''''''''''''''''''''''''' 
		Environment.value("str_ScreenName") = "Carepoint >>>> HC - First Party -  Injury Screen "
		
		If DataTable("Party_Injured","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
			objBrwpage_PartyInfo1.WebEdit("Inj_Description1").Set DataTable("Inj_Description","GL-Data")
			objBrwpage_PartyInfo1.WebEdit("Inj_CauseInjury1").Set DataTable("Inj_CauseInjury","GL-Data")
			objBrwpage_PartyInfo1.WebList("Inj_Nature1").Select DataTable("Inj_Nature","GL-Data")
			objBrwpage_PartyInfo1.WebList("Inj_BodyPart1").Select  DataTable("Inj_BodyPart","GL-Data")
			objBrwpage_PartyInfo1.WebList("Inj_InitialTreatment1").Select DataTable("Inj_InitialTreatment","GL-Data")
			If DataTable("Party_Fatality","GL-Data") = "ON" Then
				objBrwpage_PartyInfo1.WebEdit("name:=.*pdtOfDeath").Set "12/10/2015"
			End If
			
			If DataTable("Inj_InitialTreatment","GL-Data") = "NO MEDICAL TREATMENT"  or DataTable("Inj_InitialTreatment","GL-Data") = "MINOR ON-SITE REMEDIES BY EMPLOYER MEDICAL STAFF" Then
				'do Nothing
			Else
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_LName").Set DataTable("Inj_Phy_LName", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_FName").Set DataTable("Inj_Phy_FName", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_MI").Set DataTable("Inj_Phy_MI", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Add1").Set DataTable("Inj_Phy_Add1", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Add2").Set DataTable("Inj_Phy_Add2", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Zip").Set DataTable("Inj_Phy_Zip", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Phone").Set DataTable("Inj_Phy_Phone", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Fax").Set DataTable("Inj_Phy_Fax", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Email").Set DataTable("Inj_Phy_Email", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Name").Set DataTable("Inj_Hosp_Name", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Add1").Set DataTable("Inj_Hosp_Add1", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Add2").Set DataTable("Inj_Hosp_Add2", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Zip").Set DataTable("Inj_Hosp_Zip", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Phone").Set DataTable("Inj_Hosp_Phone", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Fax").Set DataTable("Inj_Hosp_Fax", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Email").Set DataTable("Inj_Hosp_Email", "GL-Data")
			End if	
			objBrwpage_PartyInfo1.WebButton("Next>>").Click
		End If
		
		Environment.value("str_ScreenName") = "Carepoint >>>> HC - First Party -  Property Damage Screen "
		'''''''''''''''''''''''''''''''''''''''''' Property Damage1 '''''''''''''''''''''''''''''''''''''''''' 
		If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
			objBrwpage_PartyInfo1.WebRadioGroup("PropertyDam1_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
			objBrwpage_PartyInfo1.WebEdit("PropertyDam1_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
			objBrwpage_PartyInfo1.WebEdit("PropertyDam1_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
			objBrwpage_PartyInfo1.WebEdit("PropertyDam1_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
			objBrwpage_PartyInfo1.WebCheckBox("PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
			objBrwpage_PartyInfo1.WebButton("Next>>").Click
		End If
		If DataTable("Party_Attorney","GL-Data") = "ON" then
			call Attorney()
		End if
	
	Else
		counter = Environment.Value("counter")
		DataTable.GetSheet("GL-Data").SetCurrentRow(counter)
		If  objBrwpage_PartyInfo1.WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Exist(5) Then
			objBrwpage_PartyInfo1.WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
			Wait(1)
		End If
		objBrwpage_PartyInfo1.WebEdit("PartyInfo1_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
		objBrwpage_PartyInfo1.WebEdit("PartyInfo1_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
		objBrwpage_PartyInfo1.WebEdit("PartyInfo1_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
		objBrwpage_PartyInfo1.WebEdit("PartyInfo1_PriPhone").Set DataTable("PartyInfo_PriPhone","GL-Data")
		objBrwpage_PartyInfo1.WebEdit("PartyInfo1_Fax").Set DataTable("PartyInfo_Fax","GL-Data")
		objBrwpage_PartyInfo1.WebEdit("PartyInfo1_Email").Set DataTable("PartyInfo_Email","GL-Data")
		objBrwpage_PartyInfo1.WebEdit("PartyInfo1_DOB").Set DataTable("PartyInfo_DOB","GL-Data")
		objBrwpage_PartyInfo1.WebList("PartyInfo1_Distribution").Select DataTable("PartyInfo_Distribution","GL-Data")
		Browser("ClaimsBrowser").Sync
		objBrwpage_PartyInfo1.WebList("PartyInfo1_Gender").Select DataTable("PartyInfo_Gender","GL-Data")
		Browser("ClaimsBrowser").Sync
		objBrwpage_PartyInfo1.WebButton("Next>>").Click
		
		'''''''''''''''''''''''''''''''''''''''''' Injury  Info1 '''''''''''''''''''''''''''''''''''''''''' 
		If DataTable("Party_Injured","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
			objBrwpage_PartyInfo1.WebEdit("Inj_Description1").Set DataTable("Inj_Description","GL-Data")
			objBrwpage_PartyInfo1.WebEdit("Inj_CauseInjury1").Set DataTable("Inj_CauseInjury","GL-Data")
			objBrwpage_PartyInfo1.WebList("Inj_Nature1").Select DataTable("Inj_Nature","GL-Data")
			objBrwpage_PartyInfo1.WebList("Inj_BodyPart1").Select  DataTable("Inj_BodyPart","GL-Data")
			objBrwpage_PartyInfo1.WebList("Inj_InitialTreatment1").Select DataTable("Inj_InitialTreatment","GL-Data")
			If DataTable("Party_Fatality","GL-Data") = "ON" Then
				objBrwpage_PartyInfo1.WebEdit("name:=.*pdtOfDeath").Set "12/10/2015"
			End If
			If DataTable("Inj_InitialTreatment","GL-Data") = "NO MEDICAL TREATMENT"  or DataTable("Inj_InitialTreatment","GL-Data") = "MINOR ON-SITE REMEDIES BY EMPLOYER MEDICAL STAFF" Then
				'do nothing
			Else
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_LName").Set DataTable("Inj_Phy_LName", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_FName").Set DataTable("Inj_Phy_FName", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_MI").Set DataTable("Inj_Phy_MI", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Add1").Set DataTable("Inj_Phy_Add1", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Add2").Set DataTable("Inj_Phy_Add2", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Zip").Set DataTable("Inj_Phy_Zip", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Phone").Set DataTable("Inj_Phy_Phone", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Fax").Set DataTable("Inj_Phy_Fax", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Phy_Email").Set DataTable("Inj_Phy_Email", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Name").Set DataTable("Inj_Hosp_Name", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Add1").Set DataTable("Inj_Hosp_Add1", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Add2").Set DataTable("Inj_Hosp_Add2", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Zip").Set DataTable("Inj_Hosp_Zip", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Phone").Set DataTable("Inj_Hosp_Phone", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Fax").Set DataTable("Inj_Hosp_Fax", "GL-Data")
				objBrwpage_PartyInfo1.WebEdit("Inj_Hosp_Email").Set DataTable("Inj_Hosp_Email", "GL-Data")
			End if	
			objBrwpage_PartyInfo1.WebButton("Next>>").Click
		End If
		'''''''''''''''''''''''''''''''''''''' Property Damage1 '''''''''''''''''''''''''''''''''''''' 
		If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
			If objBrwpage_PartyInfo1.WebRadioGroup("PropertyDam1_Location").Exist Then
			   objBrwpage_PartyInfo1.WebRadioGroup("PropertyDam1_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
			End If
			objBrwpage_PartyInfo1.WebEdit("PropertyDam1_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
			objBrwpage_PartyInfo1.WebEdit("PropertyDam1_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
			objBrwpage_PartyInfo1.WebEdit("PropertyDam1_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
			objBrwpage_PartyInfo1.WebCheckBox("PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
			objBrwpage_PartyInfo1.WebButton("Next>>").Click
		End If
		If DataTable("Party_Attorney","GL-Data") = "ON" then
			Environment.value("str_ScreenName") = "Carepoint >>>> HC - First Party -  Attorney Screen "
			call Attorney()
		End if
	End If

End Function

Function PartyInfo2()

    Environment.value("str_ScreenName") = "Carepoint >>>> HC - Second Party - Party Info Screen"
	Dim objBrwpage_PartyInfo2
	set objBrwpage_PartyInfo2=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	
	If objBrwpage_PartyInfo2.WebRadioGroup("PartyInfo2_PartyAddSame_AccAdd").Exist(5) Then
		objBrwpage_PartyInfo2.WebRadioGroup("PartyInfo2_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")
	End If
	'If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
	objBrwpage_PartyInfo2.WebEdit("PartyInfo2_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
	objBrwpage_PartyInfo2.WebEdit("PartyInfo2_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
	objBrwpage_PartyInfo2.WebEdit("PartyInfo2_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
	objBrwpage_PartyInfo2.WebEdit("PartyInfo2_PriPhone").Set DataTable("PartyInfo_PriPhone","GL-Data")
	'End if
	objBrwpage_PartyInfo2.WebEdit("PartyInfo2_Fax").Set DataTable("PartyInfo_Fax","GL-Data")
	objBrwpage_PartyInfo2.WebEdit("PartyInfo2_Email").Set DataTable("PartyInfo_Email","GL-Data")
	objBrwpage_PartyInfo2.WebEdit("PartyInfo2_DOB").Set DataTable("PartyInfo_DOB","GL-Data")
	objBrwpage_PartyInfo2.WebList("PartyInfo2_Distribution").Select DataTable("PartyInfo_Distribution","GL-Data")
	Browser("ClaimsBrowser").Sync
	objBrwpage_PartyInfo2.WebList("PartyInfo2_Gender").Select DataTable("PartyInfo_Gender","GL-Data")
	Browser("ClaimsBrowser").Sync
	Wait(1)
	objBrwpage_PartyInfo2.WebButton("Next>>").Click
	''''''''''''''''''''''''''''''''''''''' Injury 2 ''''''''''''''''''''''''''''''''''''''
	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Second Party - Injury Screen"
	If DataTable("Party_Injured","GL-Data") = "ON" Then
		Browser("ClaimsBrowser").Sync
		objBrwpage_PartyInfo2.WebEdit("Inj_Description2").Set DataTable("Inj_Description","GL-Data")
		objBrwpage_PartyInfo2.WebEdit("Inj_CauseInjury2").Set DataTable("Inj_CauseInjury","GL-Data")
		objBrwpage_PartyInfo2.WebList("Inj_Nature2").Select DataTable("Inj_Nature","GL-Data")
		objBrwpage_PartyInfo2.WebList("Inj_BodyPart2").Select  DataTable("Inj_BodyPart","GL-Data")
		objBrwpage_PartyInfo2.WebList("Inj_InitialTreatment2").Select DataTable("Inj_InitialTreatment","GL-Data")
		If DataTable("Party_Fatality","GL-Data") = "ON" Then
				objBrwpage_PartyInfo2.WebEdit("name:=.*pdtOfDeath").Set "12/10/2015"
		End If
		objBrwpage_PartyInfo2.WebButton("Next>>").Click
	End  if
	''''''''''''''''''''''''''''''''''''''' Property Damage 2 ''''''''''''''''''''''''''''''''''''''
	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Second Party - Party Damage Screen"	
	If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
		Browser("ClaimsBrowser").Sync
		objBrwpage_PartyInfo2.WebRadioGroup("PropertyDam2_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
		objBrwpage_PartyInfo2.WebEdit("PropertyDam2_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
		objBrwpage_PartyInfo2.WebEdit("PropertyDam2_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
		objBrwpage_PartyInfo2.WebEdit("PropertyDam2_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
		objBrwpage_PartyInfo2.WebButton("Next>>").Click
	End if
	If DataTable("Party_Attorney","GL-Data") = "ON" then
			Environment.value("str_ScreenName") = "Carepoint >>>> HC - Second Party -  Attorney Screen "
			call Attorney()
	End If

End Function

Function PartyInfo3()

    Environment.value("str_ScreenName") = "Carepoint >>>> HC - Third Party - Party Info "
	Dim objBrwpage_PartyInfo3
	set objBrwpage_PartyInfo3=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	
	If  objBrwpage_PartyInfo3.WebRadioGroup("PartyInfo3_PartyAddSame_AccAdd").Exist(5) Then
		objBrwpage_PartyInfo3.WebRadioGroup("PartyInfo3_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
	End If
	objBrwpage_PartyInfo3.WebEdit("PartyInfo3_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
	objBrwpage_PartyInfo3.WebEdit("PartyInfo3_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
	objBrwpage_PartyInfo3.WebEdit("PartyInfo3_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
	objBrwpage_PartyInfo3.WebButton("Next>>").Click
	''''''''''''''''''''''''''''''''' Injury 3 ''''''''''''''''''''''''''''''''' 
	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Third Party - Injury Screen"
	If DataTable("Party_Injured","GL-Data") = "ON" Then
		Browser("ClaimsBrowser").Sync
		objBrwpage_PartyInfo3.WebEdit("Inj_Description3").Set DataTable("Inj_Description","GL-Data")
		objBrwpage_PartyInfo3.WebEdit("Inj_CauseInjury3").Set DataTable("Inj_CauseInjury","GL-Data")
		objBrwpage_PartyInfo3.WebList("Inj_Nature3").Select DataTable("Inj_Nature","GL-Data")
		objBrwpage_PartyInfo3.WebList("Inj_BodyPart3").Select  DataTable("Inj_BodyPart","GL-Data")
		objBrwpage_PartyInfo3.WebList("Inj_InitialTreatment3").Select DataTable("Inj_InitialTreatment","GL-Data")
		If DataTable("Party_Fatality","GL-Data") = "ON" Then
				objBrwpage_PartyInfo3.WebEdit("name:=.*pdtOfDeath").Set "12/10/2015"
		End If
		objBrwpage_PartyInfo3.WebButton("Next>>").Click
	End if
	''''''''''''''''''''''''''''''''' Property Damage 3  ''''''''''''''''''''''''''''''''' 
	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Third Party - PD"
	If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
		Browser("ClaimsBrowser").Sync
		objBrwpage_PartyInfo3.WebRadioGroup("PropertyDam3_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
		objBrwpage_PartyInfo3.WebEdit("PropertyDam3_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
		objBrwpage_PartyInfo3.WebEdit("PropertyDam3_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
		objBrwpage_PartyInfo3.WebEdit("PropertyDam3_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
		objBrwpage_PartyInfo3.WebButton("Next>>").Click
	End If 
	
	If DataTable("Party_Attorney","GL-Data") = "ON" then
			Environment.value("str_ScreenName") = "Carepoint >>>> HC - Third Party -  Attorney Screen "
			call Attorney()
	End If
		
End function 

Function Attorney()

	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Attorney Screen"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("AttorneyList").Select DataTable("Attorney_List","GL-Data")
	Wait(3)
	If DataTable("Attorney_List","GL-Data") = "Yes" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_FirmName").Set DataTable("Attorney_FirmName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_FirstName").Set DataTable("Attorney_FirstName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_LastName").Set DataTable("Attorney_LastName","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_Address1").Set DataTable("Attorney_Address1","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_ZIP").Set DataTable("Attorney_ZIP","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_Email").Set DataTable("Attorney_Email","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_Fax").Set DataTable("Attorney_Fax","GL-Data")
	End If
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	
	
End Function


Function Additional_Information()

	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Additional Information Screen"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		
End Function

Function Assignment()

	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Assignment Screen"

	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Acc_code").Exist(5) Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("Acc_code").Select "#01" ''"Misc: Unclassified liability loss not otherwise listed"
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Run Assignment").Click
	End If 
	Wait(3)
	''''''''************************************ start: Reassign Office  ****************************************************** 
 	If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*").WebElement("innertext:=Before registering the claim, manually assign to Office 41.","html tag:=LABEL").Exist(5) Then 
 		Environment.value("str_ScreenName") = "Carepoint >>>> HC - Reassign Office"
		Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*").WebButton("name:=Reassign Office").Click
		Browser("name:=Srchssignment").Page("title:=Srchssignment").WebEdit("name:=.*PTempAssignmentPage.*pTargetCode").Set "41"
		Browser("name:=Srchssignment").Page("title:=Srchssignment").WebButton("name:=Search").Click
		Set obj = Browser("name:=Srchssignment").Page("micClass:=Page")
		Set objWebElement =  obj.webtable("column names:=Assignment;Kind;Name;Name1;Code").ChildItem(2,0,"webelement",0)
		Setting.WebPackage("ReplayType") = 2
		objWebElement.FireEvent "onclick",,,micLeftBtn 
		Setting.WebPackage("ReplayType") = 1 
		Browser("name:=Srchssignment").Page("title:=Srchssignment").WebButton("name:=Select").Click
	End If	
	''''''''************************************ END: Reassign Office  ****************************************************** 
	Environment.value("str_ScreenName") = "Carepoint >>>> HC - Assignment Screen"
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Exist Then
	   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Click
    End If	
    If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Get_Claim_Number").Exist(10) Then
       Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Get_Claim_Number").Click	
    End If
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist(7)Then 
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Click
	End IF

	WAIT(10)

End Function


Function GetClaimNumber()

	Claim_Number= Browser("ClaimsBrowser").Page("Inbox").Frame("Review_Distribution_Frame").WebTable("ClaimNumber_Table").GetROProperty("innertext")
	Claim_Number=Trim(Claim_Number)
	Claim_Number=right(Claim_Number,10)
	Environment.Value("NewClaimNumber") =  Claim_Number & "   " & Environment.Value("SCaseId")
	Print "+++++++++++++++++++++++++++++++++++++ Claim Number is +++++++++++++++ " & Environment.Value("NewClaimNumber")  & " ++++++++++++++++++++++++++++++++++++++++++++++++++++++++ "
	
	
End function

Function Review_Distribution()

		Environment.value("str_ScreenName") = "Carepoint >>>> HC - Review Distribution Screen"
		Wait(3)
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
			If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("xpath:=//button[@title='Complete']").Exist(10) Then
				Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("xpath:=//button[@title='Complete']").Click
			End If
		End If
		
End Function

 
Function Binocular_Search
 
    Dim incidentsearch,EXP_IncidentID

	If Trim(Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("Resolved-Completed").GetROProperty("innertext"))="Resolved-Completed" Then 
	Else
		EXPCase_Number=Trim(Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Link("EXP-Case").GetROProperty("text"))
	End If 
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").WebElement("Inbox").Click
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Frame").WebList("select").Select "Exception Handling"		
	Browser("ACT II").Page("get worklist for selected").WebElement("SortDate").Click
	EXP_IncidentID=Trim(Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").GetCellData(2,8))
	Set objref=createobject("Mercury.DeviceReplay")
    x=Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").ChildItem(2,3,"WebElement",0).GetRoProperty("abs_x")
    y=Browser("ACT II").Page("get worklist for selected").WebTable("WorkItem_Selection").ChildItem(2,3,"WebElement",0).GetRoProperty("abs_y")
    objref.MouseDblClick x,y,0   
    Set objref=nothing   
	
	If Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Customer Search").Exist Then		 
		Customer_Search()
	End If 

End Function 


Function Re_select_Customer()

	 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Re-select Customer").Click
	
End Function

Function Re_select_Employee()

	 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Re-select Employee").Click
	
End Function

Function Void_Incident()

	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("name:=.*PpyWorkPage.*psCaseNavigation").Select "Void"
	Wait(2)	
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebList("Enter_VodReason").Select DataTable("Enter_VodReason","GL-Data")
		If  Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("VI_Reason").Exist Then
				Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("VI_Reason").Set DataTable("VI_Reason","GL-Data")
		End If
	Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("VI_Submit").Click
	Wait (2)
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Confirm").Click

End Function


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


Function TC19_E2E_Scenario_Distribution_form_validation1()

	Browser("ClaimsBrowser").Page("Inbox").Frame("topFrame").WebElement("Search_Incident_Icon").Click
	 Browser("ACT II").Page("Claim CC Service Items").WebEdit("SI_Incident_Number").Set Environment.Value("SCaseId")
	 Browser("ACT II").Page("Claim CC Service Items").WebButton("SI_Search").Click
	Browser("ACT II").Page("Claim CC Service Items").WebElement("SI_Res_ScaseID").Click
	Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebButton("IS_Update Claim Data").Click

	If Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebButton("IS_Save").Exist Then
		ReportResult_Event micPass, "Invoking Business component: TC19_E2E_Scenario_Distribution_form_validation1" , "Page is navigated to 'General Information' Screen after 'Update Claim Data' button is clicked"
		Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").Link("IS_Site Details").Click
		Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebEdit("IS_SiteDetails_CustName").Set "Test"
	Else
	End If

	Browser("ClaimsBrowser").Page("Inbox").Frame("RoomPane").WebButton("IS_Save").Click	
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Confirm").Click

End Function



Function TC25_Close_and_Reselect_Customer_Property_Damage1()

	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("ACC_Next>>").Click
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Close").Click

	If Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").Exist then
		Browser("ClaimsBrowser").Dialog("Parent_element_Dialog").WinButton("Cancel").Click
	End if

	If Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebTable("Urgency").Exist Then
		Status = trim(Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebTable("Urgency").GetCellData(2,2))
		ScaseID = trim(Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebTable("Urgency").GetCellData(2,7))
		
			If Status = "Pending" and ScaseID = Environment.Value("SCaseId") Then
				Browser("ClaimsBrowser").Page("Inbox").Frame("myWorkList").WebElement("IB_IncidentID").Click
					Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").SetTOProperty "name","RoomPane"
			End If

	End If
End Function
 


Function TC02_E2E_scenario()

	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").Link("Incident").Click
	Wait 4
	counter = Environment.Value("counter")
	counter = counter + 1
	DataTable.GetSheet("GL-Data").SetCurrentRow(counter)

End Function


Function Set_Diaction_To_RoomPane()

   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").SetTOProperty "name","RoomPane"

End Function

Function Logout()
	
	Environment.value("str_ScreenName") = "Carepoint - HC  >>>> Logoff Screen "
	
	Browser("name:=CC.*").Page("title:=CC.*").Image("name:=Image","image type:=Image Link","Index:=0").Click
	Browser("name:=CC.*").Page("title:=CC.*").WebElement("innertext:=Log off","html id:=ItemMiddle").Click
	SystemUtil.CloseProcessByName "iexplore.exe"
	
End Function
