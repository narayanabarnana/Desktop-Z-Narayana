'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
								'Carepoint Property LOB Business Functions
								'Created By : Srirekha Talasila
								'Created On : 12/06/2016
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

ez_flag = False
firstpp_flag = False
peril_flag = False
boiler_flag = False
TPA_override = True

'Login Funtion
 Function Login()

	Environment.value("str_ScreenName") = "Carepoint >>>> Login Page "
	Systemutil.CloseProcessByName "iexplore.exe"
	Systemutil.Run "iexplore.exe",Environment.Value("CP_URL")
	Set obj_LoginPage = Browser("ClaimsBrowser").Page("LoginPage")
	obj_LoginPage.Sync
	obj_LoginPage.WebEdit("UserIdentifier").Set Environment.Value("CP_LoginId")
	obj_LoginPage.WebEdit("Password").Set Environment.Value("CP_LoginPassword") 
	obj_LoginPage.WebButton("Log In").Click
	Browser("ClaimsBrowser").Sync
	
 End function
 
'Select Property workbasket 
Function Select_Property()
	
	Environment.value("str_ScreenName") = "Carepoint >>>> Property - Select LOB "
	Browser("ClaimsBrowser").Sync
	Browser("name:=CCC.*").Page("title:=CCC.*").Link("html tag:=A","name:=My Group").Click
	Browser("name:=CCC.*").Page("title:=CCC.*").WebList("html id:=objWorkBasketSelect","html tag:=SELECT","name:=select").Select "Property"
End function

Function Select_Workitem()

	Environment.value("str_ScreenName") = "Carepoint >>>> Property - Select WorkItem "
	Wait(3)
	Browser("Customer_Browser").highlight
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

	Environment.value("str_ScreenName") = "Carepoint >>>> Property - Customer Search "
	Dim objBrwpage_CustomerSearch
	Browser("ClaimsBrowser").Sync
	set objBrwpage_CustomerSearch=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame") 
	objBrwpage_CustomerSearch.WebButton("Customer Search").Click 

	If (DataTable("Add_NewCustomer_Flag","Property") = "FALSE") Then
		If objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Exist Then
		   objBrwpage_CustomerSearch.WebEdit("CS_AccidentDate").Set DataTable("CS_Accident_Date","Property")	
		End If		
		objBrwpage_CustomerSearch.WebEdit("CS_CustomerName").Set DataTable("CS_Customer_Name","Property")
		objBrwpage_CustomerSearch.WebEdit("CS_SiteCode").Set DataTable("CS_SiteCode","Property")
		objBrwpage_CustomerSearch.WebButton("CS_Search").Click		
		Wait(5)
		Index=1
		while index<>0'''Here the condition will waits till Web Table load
			If (objBrwpage_CustomerSearch.webelement("CS_No_Matching_Data").Exist ) Then
				index=0
			Else
				index=0
				Set obj_BusinessUnit=Browser("CreationTime:=0").Page("title:=.*").Frame("name:=actionIFrame").WebTable("column names:=Click to sortBusiness Unit ,;Click to sortCustomer Name ,;Click to sortEntity Name ,;Click to sortSite Name ,;Click to sortSite Code ,;Click to sortAddress 1 ,;Click to sortAddress 2 ,;Click to sortCity ,;Click to sortState ,;Click to sortZip Code ,;Click to sortPhone ,;Click to sortFax ,","index:=23").ChildItem(2,1,"WebElement",0)''@DP
				obj_BusinessUnit.click '''This will target first row in the Customer SEarch result 
				wait(3)
				 
				objBrwpage_CustomerSearch.WebButton("html id:=submitButton").Click
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
	ElseIf DataTable("Add_NewCustomer_Flag","Property") = "TRUE" Then		
		Add_NewCustomer()
	Else
		'Do Nothing
	End If
	
End Function



Function Add_NewCustomer()

		Environment.value("str_ScreenName") = "Carepoint - Property >>>> Add New Customer "
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Add_New_Customer").Click
		Browser("ClaimsBrowser").Sync
        Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Name").Set DataTable("AddCust_Name","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Address").Set DataTable("AddCust_Address","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Zip").Set DataTable("AddCust_Zip","Property")
		Wait(1)
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Phone").Set DataTable("AddCust_Phone","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Email").Set DataTable("AddCust_Email","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("html id:=taxId").Set DataTable("AddCust_EmpTaxId","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("AddCust_Submit").Click
		wait(2)
		Browser("ClaimsBrowser").Sync
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("html id:=startProcessButton","title:=Complete this assignment","name:=.*Start Process.*").Click
		Browser("ClaimsBrowser").Sync
End Function

Function Extract_SCaseId ()

	SCase_Id=""
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("SCaseId").Exist Then
		SCase_Id = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("SCaseId").GetROProperty ("innertext")
		Print " ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ " & SCase_Id & " +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
	End If
	Environment.Value("SCaseId") = SCase_Id 

End Function

Function Incident()

		Environment.value("str_ScreenName") = "Carepoint - Property >>>> Incident Screen "
	
	    If len(Trim(Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Site_TIN").GetROProperty("value")))<9 Then 
		   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Site_TIN").Set ""
	    End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("IN_AccidentDate").Set DataTable("CS_Accident_Date","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("AccidentState").Select DataTable("IN_AccidentState","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("AccidentTime1").Select DataTable("IN_AccidentTime1","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("AccidentTime2").Select DataTable("IN_AccidentTime2","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("AccidentTime3").Select DataTable("IN_AccidentTime3","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("AccidentDescription").Set DataTable("IN_AccDescription","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("IN_ClaimSubtype").Select DataTable("IN_ClaimSubType","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("Series5OverrideReq").Set DataTable("IN_Series5OverrideReq","Property")
		If DataTable("IN_Series6OverrideReq","Property") = "ON" Then
			Browser("name:=CCC.*").Page("title:=CCC.*").Frame("name:=PegaGadget.*","Index:=1").WebCheckBox("html id:=isSeriesOverride6").Set "ON"
		End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Browser("ClaimsBrowser").Sync
		'If Duplicate Claim Exists
		If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist(5) Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Click
			Browser("ClaimsBrowser").Sync
		Else 
		'Do Nothing
		End If
		
End Function


'Created By :-  Srirekha Talasila
Function Validate_Apostrophe()

	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Policy Screen "
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("name:= << Back").Click
	Browser("ClaimsBrowser").Sync
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist(5) Then
			Environment.value("str_ScreenName") = "Carepoint - Property >>>> Duplicate Search "
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("name:= << Back").Click
			Browser("ClaimsBrowser").Sync
	End If 
	
	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Incident Screen "
	
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("IN_ClaimSubtype").Click
	
	Acc_Description = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("AccidentDescription").GetROProperty("value")
	If Acc_Description = DataTable("IN_AccDescription","Property") Then
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"Verify Apostrophe","PASS","Apostrophe is accepting in the CP Accident Description.Description is -  " & Acc_Description )	
	Else	
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"Verify Apostrophe","FAIL","Apostrophe is NOT accepting in the CP Accident Description.Description is -  " & Acc_Description )	
	End If
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	'If Duplicate Claim Exists
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist(5) Then
		Environment.value("str_ScreenName") = "Carepoint - Property >>>> Duplicate Search "
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Click
		Browser("ClaimsBrowser").Sync
	Else 
		'Do Nothing
	End If
	
End Function

'Created By :-  Srirekha Talasila
Function Validate_9_Digit_ZIP_SiteDetails()

	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Incident Screen "
	Browser("ClaimsBrowser").Sync
	Set WshShell = CreateObject("WScript.Shell")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=0").Set "123456"
	WshShell.SendKeys "{TAB}"
	
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("class:=iconError","outerhtml:=.*Please enter a valid 5 digit or 9 digit zip code.*").Exist(5) Then
		Call fn_UpdateTestResults(Environment("str_ScreenName")," VERIFY ZIP ERROR ","PASS","Validation message is displying when we enter wrong ZIP Code")	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName")," VERIFY ZIP ERROR ","FAIL","Validation message is NOT displying when we enter wrong ZIP Code")			
	End If
	
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=0").Set "123456789"
	WshShell.SendKeys "{TAB}"
	
	If  (NOT Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("class:=iconError","outerhtml:=.*Please enter a valid 5 digit or 9 digit zip code.*").Exist(5)) and (Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=0").GetROProperty("value") = "12345-6789") Then
		Call fn_UpdateTestResults(Environment("str_ScreenName")," VERIFY 9 digit ZIP ","PASS","System is Allowing 9 digit ZIP without any error.ZIP Code entered is - " & Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=0").GetROProperty("value"))	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName")," VERIFY 9 digit ZIP ","FAIL","System is NOT Allowing 9 digit ZIP without any error.ZIP Code entered is - " & Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=0").GetROProperty("value"))	
	End If
	Set WshShell = Nothing
	
End Function

'Created By :-  Srirekha Talasila
Function Validate_9_Digit_ZIP_AccidentDetails()

	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Accident Screen "
	Browser("ClaimsBrowser").Sync
	Set WshShell = CreateObject("WScript.Shell")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=1").Set "123456"
	WshShell.SendKeys "{TAB}"
	
	If  Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("class:=iconError","outerhtml:=.*Please enter a valid 5 digit or 9 digit zip code.*").Exist(5) Then
		Call fn_UpdateTestResults(Environment("str_ScreenName")," VERIFY ZIP ERROR ","PASS","Validation message is displying when we enter wrong ZIP Code")	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName")," VERIFY ZIP ERROR ","FAIL","Validation message is NOT displying when we enter wrong ZIP Code")			
	End If
	
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=1").Set "123456789"
	WshShell.SendKeys "{TAB}"
	
	If  (NOT Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("class:=iconError","outerhtml:=.*Please enter a valid 5 digit or 9 digit zip code.*").Exist(5)) and (Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=0").GetROProperty("value") = "12345-6789") Then
		Call fn_UpdateTestResults(Environment("str_ScreenName")," VERIFY 9 digit ZIP ","PASS","System is Allowing 9 digit ZIP without any error.ZIP Code entered is - " & Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=0").GetROProperty("value"))	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName")," VERIFY 9 digit ZIP ","FAIL","System is NOT Allowing 9 digit ZIP without any error.ZIP Code entered is - " & Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=0").GetROProperty("value"))	
	End If
	Set WshShell = Nothing
	
End Function

'Created By :-  Srirekha Talasila
Function Verify_City_Length()
	
	Environment.value("str_ScreenName") = "Carepoint - Property >>>>  Site Details - Verify City Length"
	Set GL_Incident=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	Browser("ClaimsBrowser").Sync
	Set WshShell = CreateObject("WScript.Shell")
	
	GL_Incident.WebEdit("name:=.*CustSiteLocation.*postalCode").Set ""
	WshShell.SendKeys "{TAB}"
	
	GL_Incident.WebEdit("name:=.*CustSiteLocation.*city","Index:=0").Set ""
	WshShell.SendKeys "{TAB}"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*CustSiteLocation.*city","Index:=0").Click
	wait(2)
	WshShell.SendKeys "ABCDEFGHIJKLMNOPQRSTUV"
	Call fn_UpdateTestResults(Environment("str_ScreenName"),"Set or Select","PASS","ABCDEFGHIJKLMNOPQRSTUV Value entered in Site City Field")	
	WshShell.SendKeys "{TAB}"
	wait(2)
	trunc_value = GL_Incident.WebEdit("name:=.*CustSiteLocation.*city").GetROProperty("value")
	
	If TRIM(trunc_value) = "ABCDEFGHIJKLMNOPQRS" Then
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"CITY NAME TRUNCATION ","PASS","CITY Value is truncated to 19 Characters when we enter more than 19.City Value after Truncation is " & trunc_value)	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"CITY NAME TRUNCATION ","FAIL","CITY Value is NOT truncated to 19 Characters when we enter more than 19.City Value after Truncation is " & trunc_value)	
	End If
	
	Set WshShell = Nothing
	
	GL_Incident.WebEdit("name:=.*CustSiteLocation.*postalCode").Set "12345"
	
End Function

'Created By :-  Srirekha Talasila
Function Verify_Claim_Series()

	Environment.value("str_ScreenName") = "Carepoint - Property  >>>> Review Distribution Screen "
	Browser("name:=CCC.*").Sync
	Call GetClaimNumber()
	Claim_Series = Left(TRIM(Environment.Value("Claim_Number")),1)

	If NOT Claim_Series = "6" Then
		Call fn_UpdateTestResults(Environment("str_ScreenName")," Verify Claim Series ","PASS","Claim Series is not started with 6 when series 6 Override is not Selected.Claim Series is - " & Claim_Series )	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName")," Verify Claim Series ","FAIL","Claim Series is started with 6 when series 6 Override is not Selected.Claim Series is - " & Claim_Series )
	End If	


End Function


'Created By :-  Srirekha Talasila
Function GetSiteDetails()
	
	Environment.value("str_ScreenName") = "Carepoint - Property  >>>> Get Site Details"
	
	Cust_Name = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=customerName","Index:=0").GetROProperty("value")
	Site_Name = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=siteName","Index:=0").GetROProperty("value")
	Environment.Value("Site_Name_CP") = Site_Name
	
	If Not Cust_Name = Site_Name Then
		Call fn_UpdateTestResults(Environment("str_ScreenName")," Customer-Site Name ","PASS","Customer and Site Name both are different.Those are  " & Cust_Name & " " & Site_Name )	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName")," Customer-Site Name ","FAIL","Customer and Site Name both are same.Those are  " & Cust_Name & " " & Site_Name &"  Select Different Customer to check the required QC" )
	End If
	
	
End Function


'Created By :-  Srirekha Talasila
Function Verify_SiteName_eZAccess()
	
		Environment.value("str_ScreenName") = "eZAccess >>>> Login Page"
		SystemUtil.Run "iexplore.exe", Environment.Value("EZ_URL")
		Wait(3)
		'EZAccess Login
		Set Obj_eZAccessBrowserPage = Browser("name:=eZAccess.*").Page("title:=eZAccess.*")
		Browser("name:=eZAccess.*").Sync
		Obj_eZAccessBrowserPage.WebEdit("name:=USER").Set Environment.Value("EZ_LoginId")
		Obj_eZAccessBrowserPage.WebEdit("name:=PASSWORD").Set Environment.Value("EZ_LoginPassword")
		Obj_eZAccessBrowserPage.WebButton("name:=Sign In").Click
		wait(2)
		Browser("name:=eZACCESS Start Page.*").Sync
'		'Open eZAccess QA Link and Maximize the Browser		
		Set Obj_eZAccessStartPage = Browser("name:=eZACCESS Start Page.*").Page("title:=eZACCESS Start Page")
		Obj_eZAccessStartPage.Link("name:=Cleanup Sessions","text:=Cleanup Sessions").Click
		Browser("name:=eZACCESS Start Page.*").Dialog("text:=Message from webpage").WinButton("text:=OK","regexpwndtitle:=OK").Click
		Wait(2)
		'Clicking on Production Sysyem link
		Obj_eZAccessStartPage.Link("name:=eZACCESS Production System","text:=eZACCESS Production System").Click
		Browser("name:=Claims Knowledge Center").Close
		Environment.value("str_ScreenName") = "eZAccess >>>> Claim Search"
		'Clicking on Claims Search 
		Browser("name:=ACT II.*").Sync
		Set Obj_eZAccessACTIIPage =Browser("name:=ACT II.*").Page("title:=ACT II")
		Browser("title:=ACT II - Internet Explorer").Sync
		Setting.WebPackage("ReplayType") = 2
		Obj_eZAccessACTIIPage.Frame("name:=topFrame","html tag:=FRAME").WebElement("innertext:=eZACCESS","html id:=ezaMenu_ezaccess").FireEvent "onmouseover"
		Setting.WebPackage("ReplayType") = 1
		Obj_eZAccessACTIIPage.Frame("name:=topFrame","html tag:=FRAME").WebElement("innertext:=Claim Search").Click
		Browser("name:=ACT II.*").Sync
		wait(2)
	    Obj_eZAccessACTIIPage.Frame("name:=sidebarFrame").WebEdit("name:=claimId").Set Environment.Value("Claim_Number")
		Obj_eZAccessACTIIPage.Frame("name:=sidebarFrame").WebButton("name:=Go!","value:=Go!").Click
		Wait(3)
		Browser("name:=ACT II.*").Sync
		InsuredName_eZA = Obj_eZAccessACTIIPage.Frame("name:=contentFrame").WebEdit("name:=localClaimSummaryAggregator.claimSummary.insdNm.lstNm").GetROProperty("value")
		
		If Environment.Value("Site_Name_CP")=Trim(InsuredName_eZA)  Then
			Call fn_UpdateTestResults(Environment("str_ScreenName")," Customer-Site-Insured Name ","PASS","Site Name passed to eZAccess.Insured name in eZaccess is " &InsuredName_eZA )	
		else
			Call fn_UpdateTestResults(Environment("str_ScreenName")," Customer-Site-Insured Name ","FAIL","Site Name NOT passed to eZAccess.Insured name in eZaccess is " &InsuredName_eZA )	
		End If
		Obj_eZAccessACTIIPage.Frame("name:=topFrame","html tag:=FRAME").WebElement("html id:=ezaMenu_logoff").Click
		
	
End Function



Function PolicySearch()
	
		Environment.value("str_ScreenName") = "Carepoint - Property >>>> Policy Screen "
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("Indeterminate").Select "SearchResults"
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PS_Policynum").Set DataTable("CS_Policynum","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Policy_Retrieve").Click		
		Wait(5)
		Cell_data = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebTable("Policy_Table").GetCellData(2,1)
		If cell_data = "" Then
			Set polobj = browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebTable("Policy_Table")
			Set polobj2 = polobj.ChildItem(2,1,"WebRadioGroup",0)
			d = polobj2.getroproperty("class")
			If d = "Radio lvInputSelection" Then
				wait(2)
		        Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("Policy_RadioButton").Click
		        wait(1)
			Else
			' Do Nothing
			End if
		End if
		If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("NoMatchingData").Exist(5) Then 
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("Indeterminate").Select "Indeterminate"
		End If 
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Browser("ClaimsBrowser").Sync
	
End Function

Function Override_TPA()
	
	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Override TPA "
	Set Obj_TPAButton = Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("actionIFrame").WebButton("name:= Override TPA","innertext:= Override TPA")
	If Obj_TPAButton.Exist(5) then
		Obj_TPAButton.Click
	Else
		'Do Nothing
	End If
	
End Function



Function Contact_Info()

	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Contact Info Screen "
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("CO_Rep_Name").Set DataTable("CO_Rep_Name","Property")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("CO_Rep_Email").Set DataTable("CO_Rep_Email","Property")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("CO_Rep_Phone").Set DataTable("CO_Rep_Phone","Property")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("CO_Report_Relation").Select DataTable("CO_Rep_Relation","Property")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("CO_CusCon_Email").Set DataTable("CO_CusCon_Email","Property")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("CO_CusCon_Fax").Set DataTable("CO_CusCon_Fax","Property")		
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
		
End function


Function Accident_Page()
	
		Environment.value("str_ScreenName") = "Carepoint - Property >>>> Accident Screen "
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Click
		Browser("ClaimsBrowser").Sync
		If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AccCode").Exist(5) then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AccCode").WaitProperty "disabled","0",10000
			x=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AccCode").getroproperty("abs_x")
			y=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AccCode").getroproperty("abs_y")
			Set objref = createobject("Mercury.DeviceReplay")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AccCode").click
			objref.MouseClick x,y,0
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AccCode").Select DataTable("ACC_AccCode","Property")
			Set objref = nothing
		End if
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AgentLoss").WaitProperty "Visible","True",1000
		x=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AgentLoss").getroproperty("abs_x")
		y=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AgentLoss").getroproperty("abs_y")
		Set objref = createobject("Mercury.DeviceReplay")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AgentLoss").click
		objref.MouseClick x,y,0
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AgentLoss").Select DataTable("ACC_AgentLoss","Property")
		Set objref = nothing

		x=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_LossLoc").getroproperty("abs_x")
		y=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_LossLoc").getroproperty("abs_y")
		Set objref = createobject("Mercury.DeviceReplay")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_LossLoc").click
		objref.MouseClick x,y,0
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_LossLoc").Select DataTable("ACC_LossLoc","Property")
		Set objref = nothing

		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_SiteAddr").Select DataTable("ACC_SiteAddress","Property")
		Rem added belwo code to enter the addredd and address2 and zip beacuse one of the scenarios though it selected yes address is not populated
		Val=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_AccZip").GetROProperty("value")
		If  Val=""   Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_AccAddress1").Set DataTable("ACC_AccAddress1","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_AccAddress2").Set DataTable("ACC_AccAddress2","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_AccZip").Set DataTable("ACC_AccZip","Property")
		End If
		Accident_SiteAddr = DataTable("ACC_SiteAddress","Property")
		If  ( Accident_SiteAddr = "No") Then
			wait(2)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_AccAddress1").Set DataTable("ACC_AccAddress1","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_AccAddress2").Set DataTable("ACC_AccAddress2","Property")
			wait(2)
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_AccZip").Set DataTable("ACC_AccZip","Property")
		Else
			'Do Nothing
		End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Comments").WaitProperty "Visible","True",1000
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Comments").Click 
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Comments").Set DataTable("ACC_Comments","Property")
		' POLICE
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("ACC_Police").Set DataTable("ACC_Police","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("ACC_Fire").Set DataTable("ACC_Fire","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("ACC_Ambulance").Set DataTable("ACC_Ambulance","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("ACC_Other").Set DataTable("ACC_Other","Property")
		If DataTable("ACC_Police","Property") = "ON" Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Pol_AuthName").Set DataTable("ACC_Pol_AuthName","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Pol_OffName").Set DataTable("ACC_Pol_OffName","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Pol_OffBatch").Set DataTable("ACC_Pol_OffBatch","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Pol_Report").Set DataTable("ACC_Pol_Report","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Pol_OffPhone").Set DataTable("ACC_Pol_OffPhone","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Pol_NCIC").Set DataTable("ACC_Pol_NCIC","Property")
		ElseIf ((DataTable("ACC_Fire","Property") = "ON") OR (DataTable("ACC_Ambulance","Property") = "ON") OR (DataTable("ACC_Other","Property") = "ON")) Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Ambu_AuthName").Set DataTable("ACC_Ambu_AuthName","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("ACC_Ambu_Report").Set DataTable("ACC_Ambu_Report","Property")
		End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click 
		Browser("ClaimsBrowser").Sync	
		
End function

'#####################################################################################################################

Function Property_Owner()

	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Property Owner Screen "
	If DataTable("PropOwn_SameasCust","Property")="ON" Then
		If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pInsuredProperty.*pOwner.*pOwnerSameAsCustomer").GetROProperty("Value") <> "true"Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pInsuredProperty.*pOwner.*pOwnerSameAsCustomer").click 
		End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Browser("ClaimsBrowser").Sync
	End if
	If DataTable("PropOwn_SameasCust","Property")="OFF" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("PO_Same").Set DataTable("PropOwn_SameasCust","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PO_Fname").Set DataTable("PropOwn_Fname","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PO_MI").Set DataTable("PropOwn_MI","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PO_Lname").Set DataTable("PropOwn_Lname","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PO_ADD1").Set DataTable("PropOwn_Add1","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PO_Zip").Set DataTable("PropOwn_Zip","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PO_PriPhone").Set DataTable("PropOwn_Phone","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PO_Fax").Set DataTable("PropOwn_Fax","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PO_Email").Set DataTable("PropOwn_Email","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PO_Distribution").Select DataTable("PropOwn_Distribution","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Browser("ClaimsBrowser").Sync
	End If
End function

Function Property_Damage()

	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Property Damage Screen "
	If Datatable("IN_ClaimSubType","Property") <> "Inland Marine" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_OwnBuilding").Select DataTable("PropDam_OwnBuilding","Property")
		If DataTable("PropDam_OwnBuilding","Property")= "Tenant" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_OwnerInformed").Select DataTable("PropDam_OwnerInformed","Property")
		End If
	End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Operration_Suspended").Select DataTable("PropDam_Operation_Suspended","Property")
	If DataTable("PropDam_Operation_Suspended","Property")="Yes" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_HowLong").Select DataTable("PropDam_HowLong","Property")
	End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Description").Set DataTable("PropDam_Description","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Location").Select DataTable("PropDam_Location","Property")
	If DataTable("PropDam_Location","Property")="Other" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_LocProp_Add1").Set DataTable("PropDam_LocProp_Add1","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_LocProp_Zip").Set DataTable("PropDam_LocProp_Zip","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_LocProp_Phone").Set DataTable("PropDam_LocProp_Phone","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_LocProp_Email").Set DataTable("PropDam_LocProp_Email","Property")
	End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Additional_Insurance").Select DataTable("PropDam_Additional_Insurance","Property")
	If  DataTable("PropDam_Additional_Insurance","Property")="Yes" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_AddInsu_CompName").Select DataTable("PropDam_AddInsu_CompName","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_AddInsu_ClaimLoss").Select DataTable("PropDam_AddInsu_ClaimLoss","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_AddInsu_Email").Select DataTable("PropDam_AddInsu_Email","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_AddInsu_Fax").Select DataTable("PropDam_AddInsu_Fax","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_AddInsu_Phone").Select DataTable("PropDam_AddInsu_Phone","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_AddInsu_Policy").Select DataTable("PropDam_AddInsu_Policy","Property")
	End If
	If DataTable("IN_ClaimSubType","Property") = "Boiler & Machinery"  Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_KindofBoiler").Select DataTable("PropDam_KindofBoiler","Property")
	End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Peril").Select DataTable("PropDam_Peril","Property")
	If  DataTable("PropDam_Peril","Property") = "Collapse"  Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Collapse_Cause").Set DataTable("PropDam_Collapse_Cause","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Collapse_DesDam").Set DataTable("PropDam_Collapse_DesDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Collapse_Discover").Set DataTable("PropDam_Collapse_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Collapse_ExtentofDam").Set DataTable("PropDam_Collapse_ExtentofDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Collapse_LossAmount").Set DataTable("PropDam_Collapse_LossAmount","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Collapse_StructureCollapse").Set DataTable("PropDam_Collapse_StructureCollapse","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Collapse_Injuries").Select DataTable("PropDam_Collapse_Injuries","Property")
		If DataTable("PropDam_Collapse_Injuries","Property") = "Yes" then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Collapse_HasWC").Select DataTable("PropDam_Collapse_HasWC","Property")
		End if
	End if
	if DataTable("PropDam_Peril","Property") = "Earthquake" then
		'Earthquake
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Earthquake_DesDam").Set DataTable("PropDam_Earthquake_DesDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Earthquake_Discover").Set DataTable("PropDam_Earthquake_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Earthquake_ExtentofDam").Set DataTable("PropDam_Earthquake_ExtentofDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Earthquake_LossAmount").Set DataTable("PropDam_Earthquake_LossAmount","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Earthquake_Building_Contents_Dam").Select DataTable("PropDam_Earthquake_Building_Contents_Dam","Property")
	End if
	if DataTable("PropDam_Peril","Property") = "Explosion" then
		'Explosion
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Explosion_Cause").Set DataTable("PropDam_Explosion_Cause","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Explosion_DesDam").Set DataTable("PropDam_Explosion_DesDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Explosion_Discover").Set DataTable("PropDam_Explosion_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Explosion_HazardMat").Set DataTable("PropDam_Explosion_HazardMat","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Explosion_LossAmount").Set DataTable("PropDam_Explosion_LossAmount","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Explosion_Injury").Select DataTable("PropDam_Explosion_Injury","Property")
	If DataTable("PropDam_Explosion_Injury","Property") = "Yes" then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Explosion_HasWC_2").Select DataTable("PropDam_Explosion_HasWC_2","Property")
	End if
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Explosion_Premises").Select DataTable("PropDam_Explosion_Premises","Property")
		If DataTable("PropDam_Explosion_Premises","Property")="Off Premises" Then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Explosion_DistofExp").Select DataTable("PropDam_Explosion_DistofExp","Property")
		End If
	End if
	if DataTable("PropDam_Peril","Property") = "Fire"  then
		'Fire
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Fire_Authority").Set DataTable("PropDam_Fire_Authority","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Fire_DesDam").Set DataTable("PropDam_Fire_DesDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Fire_Discover").Set DataTable("PropDam_Fire_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Fire_LossAmount").Set DataTable("PropDam_Fire_LossAmount","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Fire_Injury").Select DataTable("PropDam_Fire_Injury","Property")
		If DataTable("PropDam_Fire_Injury","Property")= "Yes" then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Fire_HasWC").Select DataTable("PropDam_Fire_HasWC","Property")
		End if
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Fire_PropSlavaged").Select DataTable("PropDam_Fire_PropSlavaged","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Fire_StepsTaken").Select DataTable("PropDam_Fire_StepsTaken","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Fire_StockInvent_Dam").Select DataTable("PropDam_Fire_StockInvent_Dam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Fire_TempRepair").Select DataTable("PropDam_Fire_TempRepair","Property")
		If DataTable("PropDam_Fire_TempRepair","Property") = "Yes" then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Fire_TempRepairScope").Select DataTable("PropDam_Fire_TempRepairScope","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Fire_TempRepairCost").Select DataTable("PropDam_Fire_TempRepairCost","Property")
		End if
	End if
	if DataTable("PropDam_Peril","Property") = "Lightning" then
		'Lightning
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Lightning_DamDes").Set DataTable("PropDam_Lightning_DamDes","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Lightning_Discover").Set DataTable("PropDam_Lightning_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Lightning_LossAmount").Set DataTable("PropDam_Lightning_LossAmount","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Lightning_AlarmDam").Select DataTable("PropDam_Lightning_AlarmDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Lightning_ResultingFires").Select DataTable("PropDam_Lightning_ResultingFires","Property")
	End if
	if DataTable("PropDam_Peril","Property") = "Other"  then
		'Other
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Other_DesDam").Set DataTable("PropDam_Other_DesDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Other_Discover").Set DataTable("PropDam_Other_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Other_ExtentofDam").Set DataTable("PropDam_Other_ExtentofDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Other_LossAmount").Set DataTable("PropDam_Other_LossAmount","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Other_Building_Contents_Dam").Select DataTable("PropDam_Other_Building_Contents_Dam","Property")
	End if
	if DataTable("PropDam_Peril","Property") = "Smoke"  then
		'Smoke
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Smoke_DesDam").Set DataTable("PropDam_Smoke_DesDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Smoke_Discover").Set DataTable("PropDam_Smoke_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Smoke_LossAmount").Set DataTable("PropDam_Smoke_LossAmount","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Smoke_Source").Set DataTable("PropDam_Smoke_Source","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Smoke_ExtentDam").Set DataTable("PropDam_Smoke_ExtentDam","Property")
	End if
	if DataTable("PropDam_Peril","Property") = "Theft"  then
	'Theft
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Theft_DesDam").Set DataTable("PropDam_Theft_DesDam","Property")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Theft_Discover").Set DataTable("PropDam_Theft_Discover","Property")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Theft_LossAmount").Set DataTable("PropDam_Theft_LossAmount","Property")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Theft_TypeofProp").Set DataTable("PropDam_Theft_TypeofProp","Property")
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Theft_Cause").Select DataTable("PropDam_Theft_Cause","Property")
	'If DataTable("PropDam_Theft_Cause","Property")= "Yes" then
		'Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Theft_EntryGain").Select DataTable("PropDam_Theft_EntryGain","Property")
	'End if
	End if
	if DataTable("PropDam_Peril","Property") = "Vandalism"  then
		'Vandalism
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Vandalism_DesDam").Set DataTable("PropDam_Vandalism_DesDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Vandalism_Discover").Set DataTable("PropDam_Vandalism_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Vandalism_ExtentofDam").Set DataTable("PropDam_Vandalism_ExtentofDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Vandalism_LossAmount").Set DataTable("PropDam_Vandalism_LossAmount","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Vandalism_Building_Contents_Dam").Select DataTable("PropDam_Vandalism_Building_Contents_Dam","Property")
	End if
	if DataTable("PropDam_Peril","Property") = "Vehicle"  then
		'Vehicle
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Vehicle_DesDam").Set DataTable("PropDam_Vehicle_DesDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Vehicle_Discover").Set DataTable("PropDam_Vehicle_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Vehicle_ExtentofDam").Set DataTable("PropDam_Vehicle_ExtentofDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Vehicle_LossAmount").Set DataTable("PropDam_Vehicle_LossAmount","Property")
	End if
	if DataTable("PropDam_Peril","Property") = "Water"  then	
		'Water
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Water_AreaAffected").Set DataTable("PropDam_Water_AreaAffected","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Water_Cause").Set DataTable("PropDam_Water_Cause","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Water_Depth").Set DataTable("PropDam_Water_Depth","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Water_DesDam").Set DataTable("PropDam_Water_DesDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Water_Discover").Set DataTable("PropDam_Water_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Water_LossAmount").Set DataTable("PropDam_Water_LossAmount","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Water_WaterEntered").Set DataTable("PropDam_Water_WaterEntered","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Water_Plumbing").Select DataTable("PropDam_Water_Plumbing","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Water_SignofMold").Select DataTable("PropDam_Water_SignofMold","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Water_WaterRemoved").Select DataTable("PropDam_Water_WaterRemoved","Property")
	End if
	if DataTable("PropDam_Peril","Property") = "Wind"  then
		'Wind
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Wind_BuildingSize").Set DataTable("PropDam_Wind_BuildingSize","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Wind_Discover").Set DataTable("PropDam_Wind_Discover","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Wind_ExtentofDam").Set DataTable("PropDam_Wind_ExtentofDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Wind_How_Many_Stories").Set DataTable("PropDam_Wind_How_Many_Stories","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Wind_InteriorDam").Set DataTable("PropDam_Wind_InteriorDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("PropDam_Wind_LossAmount").Set DataTable("PropDam_Wind_LossAmount","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Wind_ContentDam").Select DataTable("PropDam_Wind_ContentDam","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Wind_DamType").Select DataTable("PropDam_Wind_DamType","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Wind_RoofDam").Select DataTable("PropDam_Wind_RoofDam","Property")
		If DataTable("PropDam_Wind_RoofDam","Property") = "Yes" then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Wind_RoofSize").Select DataTable("PropDam_Wind_RoofSize","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Wind_TypeofRoof").Select DataTable("PropDam_Wind_TypeofRoof","Property")
		End if
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Wind_TempRepair").Select DataTable("PropDam_Wind_TempRepair","Property")
		If DataTable("PropDam_Wind_TempRepair","Property") = "Yes" then
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Wind_RepairScope").Select DataTable("PropDam_Wind_RepairScope","Property")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("PropDam_Wind_RepairCost").Select DataTable("PropDam_Wind_RepairCost","Property")
		End if
	End if 
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Browser("ClaimsBrowser").Sync	
		
End function



Function Witness()
		
		Environment.value("str_ScreenName") = "Carepoint - Property >>>> Witness Screen "
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("WitnessList").Select DataTable("Witness_List","Property")
	If DataTable("Witness_List","Property") = "Yes" Then	
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Wit_FirstName").Set DataTable("Witness_FirstName","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Wit_LastName").Set DataTable("Witness_LastName","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pWitness.*pwitnessDetails.*l1.*pAddressDetails.*paddressLine1","html tag:=INPUT").Set DataTable("Witness_Address1","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pWitness.*pwitnessDetails.*l1.*pAddressDetails.*paddressLine2","html tag:=INPUT").Set DataTable("Witness_Address2","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*PpyWorkPage.*pClaimData.*pNoticeDataClm.*pWitness.*pwitnessDetails.*l1.*pAddressDetails.*ppostalCode","html tag:=INPUT").Set DataTable("Witness_Zip","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=HomePhone.*","html tag:=INPUT").Set DataTable("Witness_PrimaryPhone","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=Fax.*","html tag:=INPUT").Set DataTable("Witness_Fax","Property")
	End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Browser("ClaimsBrowser").Sync
	

End Function

Function Attorney()
		
		Environment.value("str_ScreenName") = "Carepoint - Property >>>> Attorney Screen "
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("AttorneyList").Select DataTable("Attorney_List","Property")
	If DataTable("Attorney_List","Property") = "Yes" Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_FirmName").Set DataTable("Attorney_FirmName","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_FirstName").Set DataTable("Attorney_FirstName","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_LastName").Set DataTable("Attorney_LastName","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_Address1").Set DataTable("Attorney_Address1","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_ZIP").Set DataTable("Attorney_ZIP","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_Email").Set DataTable("Attorney_Email","Property")
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Att_Fax").Set DataTable("Attorney_Fax","Property")
	End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Browser("ClaimsBrowser").Sync	
End Function


Function Additional_Information()

		Environment.value("str_ScreenName") = "Carepoint - Property >>>> Additional Info Screen "
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
		Browser("ClaimsBrowser").Sync
End Function

Function QuickStartActivity()
	
	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Quick Start Activity Screen "
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("QuickStart1").Set "ON"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("QuickStart2").Set "ON"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("QuickStart3").Set "ON"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("QuickStart4").Set "ON"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("QuickStart5").Set "ON"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("QuickStart6").Set "ON"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("QuickStart7").Set "ON"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("QuickStart8").Set "ON"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebCheckBox("QuickStart9").Set "ON"
	Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync	

End Function

Function Assignment()

	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Assignment Screen "
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Run Assignment").Exist(5)Then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebList("ACC_AccCode").Select "#1"
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Run Assignment").Click
		Browser("ClaimsBrowser").Sync
	End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Ass_Save").Click
		Browser("ClaimsBrowser").Sync
	If (DataTable("AccidentCode_Override_TPA","Property") = "TRUE") Then
		Call ReassignOffice()
	End If
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("Get_Claim_Number").Click
	If   Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Exist(5) Then
		 Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("No Duplicates Found").Click
		 Browser("ClaimsBrowser").Sync
	End If
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("CallBack_Override").Exist(5) then
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("CallBack_Reason").Set "test"
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebButton("CallBack_Override").Click
	End if 

End Function

Function GetClaimNumber()
	
	Environment.value("str_ScreenName") = "Carepoint - Property >>>> Claim Number Screen "
	Claim_Number= Browser("ClaimsBrowser").Page("Inbox").Frame("name:=PegaGadget0Ifr").WebTable("innertext:=Status.*").GetROProperty("innertext")
	Claim_Number=Trim(Claim_Number)
	Claim_Number=right(Claim_Number,10)
	Environment.Value("Claim_Number") = Claim_Number
	Environment.Value("NewClaimNumber") =  Claim_Number & "    " & Environment.Value("SCaseId")
	Print " ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ "&Environment.Value("NewClaimNumber") &  "     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
	 
End Function

'Created By :-  Srirekha Talasila
'This will handle Distributions in Review Screen 

Function Review_Distribution()
	
	On Error Resume Next
	Browser("name:=CCC.*").Sync
	Call GetClaimNumber()
	Environment.value("str_ScreenName") = "Carepoint - Property  >>>> Review Distribution Screen "
	If Datatable("AccidentCode_Override_TPA","Property") = "TRUE" and TPA_override = "True" Then
			Call Override_TPA()
			TPA_override = false
	Else
		
		If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("name:=Start Process.*").Exist(5) Then
			''Log Off	
		Else
				If  DataTable("CS_Policynum","Property") = "28626530"Then
						Claim_Number =" Farmer's Policy"
						Environment.Value("NewClaimNumber") =  Claim_Number & "    " & Environment.Value("SCaseId")
						Print " ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ "&Environment.Value("NewClaimNumber") &  "     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
				Else
					Call GetClaimNumber()
				End If
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
			If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("xpath:=//button[@title='Complete']").Exist Then
				Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("xpath:=//button[@title='Complete']").Click
			End If
		End If 
	End If	

End Function
 

Function ReassignOffice()

	Environment.value("str_ScreenName") = "Carepoint - Property  >>>> Reassign Office Screen "
	
	If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*","html tag:=IFRAME").WebElement("innertext:=Because the TPA Override option was selected, please manually assign the correct Zurich handling office","html tag:=LABEL").Exist Then 
		Browser("name:=CCC.*").Page("title:=CCC.*").Frame("title:=FNOL.*").WebButton("name:=Reassign Office").Click
		Browser("name:=Srchssignment").Page("title:=Srchssignment").WebEdit("name:=.*PTempAssignmentPage.*pTargetCode").Set "41"
		Browser("name:=Srchssignment").Page("title:=Srchssignment").WebButton("name:=Search").Click
		Set obj = Browser("name:=Srchssignment").Page("micClass:=Page")
		Set objWebElement =  obj.webtable("column names:=Assignment;Kind;Name;Name1;Code").ChildItem(2,0,"webelement",0)
		Setting.WebPackage("ReplayType") = 2
		objWebElement.FireEvent "ondblclick",,,micLeftBtn 
		Setting.WebPackage("ReplayType") = 1 
		Browser("name:=Srchssignment").Page("title:=Srchssignment").WebButton("name:=Select").Click
	End If	


End Function


Function Logout()
	
	Environment.value("str_ScreenName") = "Carepoint - Property  >>>> Logoff Screen "
	
	Browser("name:=CC.*").Page("title:=CC.*").Image("name:=Image","image type:=Image Link","Index:=0").Click
	Browser("name:=CC.*").Page("title:=CC.*").WebElement("innertext:=Log off","html id:=ItemMiddle").Click
	SystemUtil.CloseProcessByName "iexplore.exe"
	
End Function

'Created By :-  Srirekha Talasila
'This will Verify Search Functionality using S-Case and Claim Number
Function Binocular_Search()

	Environment.value("str_ScreenName") = "Carepoint - Property  >>>> Binocular Search Screen "
	Browser("ClaimsBrowser").Page("Inbox").Link("Binocular_Search").Click
	Wait(3)
	Browser("SearchIncident").Sync
	Browser("SearchIncident").Page("SearchIncident").WebEdit("Claim_Number").Set DataTable("Claim_Number","Property")
	Browser("SearchIncident").Page("SearchIncident").WebButton("Search_Btn").Click
	Browser("SearchIncident").Sync
	ClaimNumber = Browser("SearchIncident").Page("SearchIncident").WebTable("Search_Results_Table").GetCellData(2,2)
	
	If TRIM(DataTable("Claim_Number","Property")) = TRIM(ClaimNumber) Then
		Set Obj_Result = Browser("SearchIncident").Page("SearchIncident").WebTable("Search_Results_Table").ChildItem(2,0,"WebElement",0)
		Obj_Result.highlight
		Obj_Result.click
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY CLAIMNUMBER","PASS","Claim Number " &ClaimNumber& " Exist in WebTable")	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY CLAIMNUMBER","FAIL","Claim Number " &ClaimNumber& " NOT Exist in WebTable")			
	End If
	Browser("SearchIncident").Page("SearchIncident").WebButton("Clear_Btn").Click
	Browser("SearchIncident").Sync
	Browser("SearchIncident").Page("SearchIncident").WebEdit("Incident_Num").Set DataTable("SCase","Property")
	Browser("SearchIncident").Page("SearchIncident").WebRadioGroup("S-Case-Include").Select "Include"
	Browser("SearchIncident").Page("SearchIncident").WebButton("Search_Btn").Click
	Browser("name:=Claim CC Service.*").Sync
	SNumber = Browser("Claim CC Service Items").Page("Claim CC Service Items").WebTable("Claim_Number").GetCellData(2,1)
	If TRIM(DataTable("SCase","Property")) = TRIM(SNumber) Then
		Set Obj_Result1 = Browser("Claim CC Service Items").Page("Claim CC Service Items").WebTable("Claim_Number").ChildItem(2,0,"WebElement",0)
		Obj_Result1.highlight
		Obj_Result1.click
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY S-CASE","PASS","S-CASE " &SNumber& " Exist in WebTable")	
	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY S-CASE","FAIL"," S-CASE " &SNumber& " NOT Exist in WebTable")			
	End If
	Browser("name:=Claim CC Service.*").Close

End Function 
