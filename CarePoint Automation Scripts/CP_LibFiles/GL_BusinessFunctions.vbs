'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
								'Carepoint GL Business Functions
								'Updated By : Srirekha Talasila
								'Updated On : 12/21/2016
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Function Login()

	Environment.value("str_ScreenName") = "Carepoint >>>> GL - Login Page "
	SystemUtil.CloseProcessByName "iexplore.exe"
	Systemutil.Run "iexplore.exe", Environment.Value("CP_URL")	
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("UserIdentifier").Set Environment.Value("CP_LoginId")
	Browser("ClaimsBrowser").Page("LoginPage").WebEdit("Password").Set Environment.Value("CP_LoginPassword") 
	Browser("ClaimsBrowser").Page("LoginPage").WebButton("Log In").Click
	Browser("ClaimsBrowser").Sync
	
 End function


Function Select_GL()
	
	Environment.value("str_ScreenName") = "Carepoint >>>> GL - Select GL "
	Browser("name:=CCC.*").Page("title:=CCC.*").Link("html tag:=A","name:=My Group").Click
	Browser("name:=CCC.*").Page("title:=CCC.*").WebList("html id:=objWorkBasketSelect","html tag:=SELECT","name:=select").Select "General Liability"
	
End function


Function Select_Workitem()

	Environment.value("str_ScreenName") = "Carepoint >>>> GL - Select WorkItem "
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
'						setting.webpackage("ReplayType")=2
'						Browser("Customer_Browser").Page("WorkList_Basket").WebTable("WorkItem_Selection").ChildItem(row,3,"WebElement",0).fireevent("Onmouseover")
'						Browser("Customer_Browser").Page("WorkList_Basket").WebTable("WorkItem_Selection").ChildItem(row,3,"WebElement",0).fireevent("OnDblClick")
'						setting.webpackage("ReplayType")=1
						Exit For				 			
	         		End If	
					If row=7 Then
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

	Environment.value("str_ScreenName") = "Carepoint >>>> GL - Customer Search "
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
	ElseIf DataTable("Add_NewCustomer_Flag","GL-Data") = "TRUE" Then		
		Add_NewCustomer()
	Else
		Employee_Search()
	End If
	
End Function


Function Add_NewCustomer()

		Environment.value("str_ScreenName") = "Carepoint - GL >>>> Add New Customer "
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebButton("Add New Customer").Click
		Browser("ClaimsBrowser").sync
		wait(1)
        Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Name").Set DataTable("AddCust_Name","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Address").Set DataTable("AddCust_Address","GL-Data")
		Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame").WebEdit("AddCust_Zip").Set DataTable("AddCust_Zip","GL-Data")
		wait(1)
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

	Environment.value("str_ScreenName") = "Carepoint - GL >>>>  Employee Search "
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
	If Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("SCaseId").Exist(2) Then
		SCase_Id = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebElement("SCaseId").GetROProperty ("innertext")
		Print "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++  " & SCase_Id & "  +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
	End If
	Environment.Value("SCaseId") = SCase_Id 

End Function


Function Incident()
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>>  Incident Screen "
	Set GL_Incident=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	Browser("ClaimsBrowser").Sync
	GL_Incident.WebEdit("IN_AccidentDate").Set DataTable("CS_Accident_Date","GL-Data")
	GL_Incident.WebList("AccidentState").Select DataTable("IN_AccidentState","GL-Data")
	GL_Incident.WebList("AccidentTime1").Select DataTable("IN_AccidentTime1","GL-Data")
	GL_Incident.WebList("AccidentTime2").Select DataTable("IN_AccidentTime2","GL-Data")
	GL_Incident.WebList("AccidentTime3").Select DataTable("IN_AccidentTime3","GL-Data")
	GL_Incident.WebEdit("AccidentDescription").Set DataTable("IN_AccDescription","GL-Data")
	GL_Incident.WebList("Catagory").Select DataTable("IN_Category","GL-Data")
	If  DataTable("IN_Category","GL-Data") = "Innkeepers/Guest Property" Then
		GL_Incident.WebCheckBox("Series6_Override").Set DataTable("IN_Series6OverrideReq","GL-Data")
	End If
	GL_Incident.WebButton("Next>>").Click
	'If Duplicate Claim Exists
	If GL_Incident.WebButton("No Duplicates Found").Exist(5) Then
		GL_Incident.WebButton("No Duplicates Found").Click
		Browser("ClaimsBrowser").Sync
	Else 
		'Do Nothing
	End If
	
End Function


Function PolicySearch()

	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Policy Search "
	
	Set GL_PolSearch=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	Browser("ClaimsBrowser").Sync

	cell_data = GL_PolSearch.WebTable("Policy_Table").GetCellData(2,1)
	If cell_data = ""  and DataTable("CS_Policynum","GL-Data") = "" Then
		Set polobj = GL_PolSearch.WebTable("Policy_Table")
		Set polobj2 = polobj.ChildItem(2,1,"WebRadioGroup",0)
		d = polobj2.getroproperty("class")
		If d = "Radio lvInputSelection" Then
			GL_PolSearch.WebRadioGroup("Policy_RadioButton").Click
			GL_PolSearch.WebButton("Next>>").Click
		Else
	    End if
	End if
	pol_flag = False
	If GL_PolSearch.WebElement("innertext:=No matching policy records found.*","innerhtml:=No matching policy records found.*").Exist(5) OR DataTable("CS_Policynum","GL-Data") <> "" Then
		pol_flag = True
		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebRadioGroup("PolicyList_Radio_Btn").Select "SearchResults"
		GL_PolSearch.WebEdit("PS_Policynum").Set (DataTable("CS_Policynum","GL-Data"))
		GL_PolSearch.WebEdit("html id:=zpsMonthsPrior").Set Trim(DataTable("MonthsPrior","GL-Data"))
		GL_PolSearch.WebButton("Policy_Retrieve").Click
		cell_data = GL_PolSearch.WebTable("Policy_Table").GetCellData(2,1)
		Wait 1
		If cell_data = ""  Then
			If GL_PolSearch.WebRadioGroup("Policy_RadioButton").Exist(5)  Then
				GL_PolSearch.WebRadioGroup("Policy_RadioButton").Click
				GL_PolSearch.WebButton("Next>>").Click
			End If
		Else 
			GL_PolSearch.WebRadioGroup("Indeterminate").Select "Indeterminate"
			GL_PolSearch.WebButton("Next>>").Click
		End if
	End If
	Browser("ClaimsBrowser").Sync
End Function


Function Override_TPA()
   	
   	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Override TPA "
   	Browser("ClaimsBrowser").Sync
   	Set GL_TPA=Browser("ClaimsBrowser").Page("Inbox").Frame("actionIFrame") 
   	If GL_TPA.WebButton("Override_TPA").Exist(5) then
		 GL_TPA.WebButton("Override_TPA").Click
		 Browser("ClaimsBrowser").Sync
	Else
		'Do Nothing
	End If
	
End Function

'Created By :-  Srirekha Talasila
Function Verify_Office_35()
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Verify Office NOT Care Center 35"
   	Browser("ClaimsBrowser").Sync
   	Set GL_BrwPage = Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
   	Office = GL_BrwPage.WebTable("name:=.*Assign.*targetState").GetCellData(2,4)
   	If Trim(Office) = Trim(" Care Center - 35 ") Then
   		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY Office ","FAIL","Claim is assigned to" & Office & "when we Select Zurich Employee as NO ")
   	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY Office","PASS","Claim is Assigned to " & Office & " when we Select Zurich Employee as NO ")	
   	End If
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Assignment "
	GL_BrwPage.WebButton("name:= << Back").Click
	Browser("ClaimsBrowser").Sync
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Additional Info "
	GL_BrwPage.WebButton("name:= << Back").Click
	Browser("ClaimsBrowser").Sync
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Witness Screen "
	GL_BrwPage.WebButton("name:= << Back").Click
	Browser("ClaimsBrowser").Sync
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Injury Screen "
	GL_BrwPage.WebButton("name:= << Back").Click
	Browser("ClaimsBrowser").Sync
	
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> First Party - Party Info Screen "
	GL_BrwPage.WebRadioGroup("PartyInfo1_ZurichEmp").Select "true"
	GL_BrwPage.WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Injury Screen "
	GL_BrwPage.WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Witness Screen "
	GL_BrwPage.WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Additional Info "
	GL_BrwPage.WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Verify Office Care Center 35"
	
	Office1 = GL_BrwPage.WebTable("name:=.*Assign.*targetState").GetCellData(2,4)
   	
   	If Trim(Office1) = Trim(" Care Center - 35 ") Then
   		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY Office","PASS","Claim is Assigned to " & Office1 & " when we Select Zurich Employee as YES ")	
   	Else
		Call fn_UpdateTestResults(Environment("str_ScreenName"),"VERIFY Office ","FAIL","Claim is not assigned to Care Center - 35 when we Select Zurich Employee as YES ")
   	End If
	
	
End Function

'Created By :-  Srirekha Talasila
Function Verify_Validation_Message_Distributions()

		Environment.value("str_ScreenName") = "Carepoint - GL  >>>> Verify Distributions Validations Message "
		On Error Resume Next
		Browser("name:=CCC.*").Sync
		
		Set Obj_Page = Browser("name:=CCC.*").Page("title:=CCC.*")
		Set obj_ActionIFrame = Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame")
		
		
		'Email	
		obj_ActionIFrame.WebButton("html id:=RLAdd","html tag:=BUTTON").Click
		wait(1)
		Call Add_New_Distribution("Email")
		Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("xpath:=//button[@title='Complete']").Click
		Browser("name:=CCC.*").Sync
		If Browser("name:=Access required").Exist(2) Then
			Browser("name:=Access required").Close
		End If
		If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebElement("class:=iconError","outerhtml:=.* Please enter email.*","Index:=0").Exist(5) Then
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"EMAIL - VERIFY Email Address Validation ","PASS","Validation message is displying when we Select Method as EMAIL with Empty Email Address")	
		Else
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"EMAIL - VERIFY Email Address Validation ","FAIL","Validation message is NOT displying when we Select Method as EMAIL with Empty Email Address ")	
		End If
		
		
		'ELECACK
		obj_ActionIFrame.WebButton("html id:=RLAdd","html tag:=BUTTON").Click
		Call Add_New_Distribution("ELECACK")
		Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("xpath:=//button[@title='Complete']").Click
		Browser("name:=CCC.*").Sync
		If Browser("name:=Access required").Exist(2) Then
			Browser("name:=Access required").Close
		End If
		If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebElement("class:=iconError","outerhtml:=.* Please enter email.*","Index:=1").Exist(5) Then
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"ELECACK - VERIFY Email Address Validation ","PASS","Validation message is displying when we Select Method as ELECACK with Empty Email Address")	
		Else
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"ELECACK - VERIFY Email Address Validation ","FAIL","Validation message is NOT displying when we Select Method as ELECACK with Empty Email Address ")	
		End If
		
		
		'Fax
		obj_ActionIFrame.WebButton("html id:=RLAdd","html tag:=BUTTON").Click
		Call Add_New_Distribution("Fax")
		Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("xpath:=//button[@title='Complete']").Click
		Browser("name:=CCC.*").Sync
		If Browser("name:=Access required").Exist(2) Then
			Browser("name:=Access required").Close
		End If
		If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebElement("class:=iconError","outerhtml:=.* Please enter fax.*","Index:=0").Exist(5) Then
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"FAX - VERIFY FAX Validation ","PASS","Validation message is displying when we Select Method as FAX with Empty FAX NO")	
		Else
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"FAX - VERIFY FAX Validation ","FAIL","Validation message is NOT displying when we Select Method as FAX with Empty FAX NO ")	
		End If
		
		
		'Mail
		obj_ActionIFrame.WebButton("html id:=RLAdd","html tag:=BUTTON").Click
		Call Add_New_Distribution("Mail")
		Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame").WebButton("xpath:=//button[@title='Complete']").Click
		If Browser("name:=Access required").Exist(2) Then
			Browser("name:=Access required").Close
		End If
		Browser("name:=CCC.*").Sync
		Set Obj_Frame = Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame")
		If Obj_Frame.WebElement("class:=iconError","outerhtml:=.* Please enter city.*","Index:=0").Exist(5) AND Obj_Frame.WebElement("class:=iconError","outerhtml:=.* Please enter zip.*","Index:=0").Exist(5) AND  Obj_Frame.WebElement("class:=iconError","outerhtml:=.* Please enter address 1.*","Index:=0").Exist(5) Then
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"MAIL - VERIFY Mail Validation ","PASS","Validation messages are displying when we Select Method as MAIL with Empty MAIL ADDRESS")	
		Else
			Call fn_UpdateTestResults(Environment("str_ScreenName"),"MAIL - VERIFY Mail Validation ","FAIL","Validation messages are  NOT displying when we Select Method as MAIL with Empty FAX ADDRESS ")	
		End If
	
	
End Function

'Created By :-  Srirekha Talasila
Function Add_New_Distribution(MethodName)
	
		Set Obj_Page = Browser("name:=CCC.*").Page("title:=CCC.*")
		Set obj_ActionIFrame1 = Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html tag:=IFRAME","name:=actionIFrame")
		'Descriptive object to identify  Web List objects in Review Distribution Screen.
		Set DList1 = Description.Create
		DList1("micclass").value="WebList" 
		Set  Obj_WebList = obj_ActionIFrame1.ChildObjects(DList1) 
		NoOfWebListObj = Obj_WebList.Count
		
		For Counter=0 to NoOfWebListObj-1 
			If Right(Obj_WebList(Counter).getroproperty("name"),15) = "pdistMethodName" then
		     		ChannelValue = Obj_WebList(Counter).getroproperty("value")
		     		CommmonValue = Left(Obj_WebList(Counter).getroproperty("name"),40)
		     		ActalValue = Obj_WebList(Counter).getroproperty("name")
					If ChannelValue = "#0" Then
						Set  DChannel=description.Create
						DChannel("micclass").value="WebList"
						DChannel("name").value= ActalValue
						DChannel("html id").value= "distMethodName"
						DChannel("name").RegularExpression = false
		     			obj_ActionIFrame1.WebList(DChannel).Select MethodName
		     			
		     			If MethodName <> "ELECACK" Then
		     				Set  DCheckbox1=description.Create
							DCheckbox1("micclass").value="WebCheckBox"
							DCheckbox1("name").value= CommmonValue & "$plossNoticeInd"
							DCheckbox1("type").value= "checkbox"
							DCheckbox1("name").RegularExpression = false
							wait(2)
							obj_ActionIFrame1.WebCheckBox(DCheckbox1).Set "ON"
		     			End If
		     			
						
		     		End IF
		     		
		     End IF
  		Next 
	
End Function


Function Contact_Info()

	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Contact Info "
	Set GL_ConInfo=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	Browser("ClaimsBrowser").Sync
	GL_ConInfo.WebEdit("CO_Rep_Name").Set DataTable("CO_Rep_Name","GL-Data")
	GL_ConInfo.WebEdit("CO_Rep_Email").Set DataTable("CO_Rep_Email","GL-Data")
	GL_ConInfo.WebEdit("CO_Rep_Phone").Set DataTable("CO_Rep_Phone","GL-Data")
	GL_ConInfo.WebList("CO_Report_Relation").Select DataTable("CO_Rep_Relation","GL-Data")
	GL_ConInfo.WebEdit("CO_CusCon_Email").Set DataTable("CO_CusCon_Email","GL-Data")
	GL_ConInfo.WebEdit("CO_CusCon_Fax").Set DataTable("CO_CusCon_Fax","GL-Data")		
	GL_ConInfo.WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	
End function


Function Accident_Page()

	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Accident Screen "
	Set GL_Accident=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	GL_Accident.WebButton("Ass_Save").Click
	Browser("ClaimsBrowser").Sync
	GL_Accident.WebList("ACC_AccCode").Select DataTable("ACC_AccCode","GL-Data")
	GL_Accident.WebList("ACC_AgentLoss").Select DataTable("ACC_AgentLoss","GL-Data")
	GL_Accident.WebList("ACC_LossLoc").Select DataTable("ACC_LossLoc","GL-Data")
	GL_Accident.WebList("ACC_SiteAddress").Select DataTable("ACC_SiteAddress","GL-Data")
	Accident_SiteAddr = DataTable("ACC_SiteAddress","GL-Data")
	If  (Accident_SiteAddr = "No") Then
		GL_Accident.WebEdit("ACC_AccAddress1").Set DataTable("ACC_AccAddress1","GL-Data")
		GL_Accident.WebEdit("ACC_AccAddress2").Set DataTable("ACC_AccAddress2","GL-Data")
		GL_Accident.WebEdit("ACC_AccZip").Set DataTable("ACC_AccZip","GL-Data")
	Else
		'Do Nothing
		If GL_Accident.WebEdit("ACC_AccZip").GetROProperty("value") = "" Then
			GL_Accident.WebEdit("ACC_AccZip").Set "12345"
		End If
	End If
	GL_Accident.WebEdit("ACC_Comments").Set DataTable("ACC_Comments","GL-Data")
	' POLICE
	GL_Accident.WebCheckBox("ACC_Police").Set DataTable("ACC_Police","GL-Data")
	GL_Accident.WebCheckBox("ACC_Fire").Set DataTable("ACC_Fire","GL-Data")
	GL_Accident.WebCheckBox("ACC_Ambulance").Set DataTable("ACC_Ambulance","GL-Data")
	GL_Accident.WebCheckBox("ACC_Other").Set DataTable("ACC_Other","GL-Data")
	If DataTable("ACC_Police","GL-Data") = "ON" Then
		GL_Accident.WebEdit("ACC_Pol_AuthName").Set DataTable("ACC_Pol_AuthName","GL-Data")
		GL_Accident.WebEdit("ACC_Pol_OffName").Set DataTable("ACC_Pol_OffName","GL-Data")
		GL_Accident.WebEdit("ACC_Pol_OffBatch").Set DataTable("ACC_Pol_OffBatch","GL-Data")
		GL_Accident.WebEdit("ACC_Pol_Report").Set DataTable("ACC_Pol_Report","GL-Data")
		GL_Accident.WebEdit("ACC_Pol_OffPhone").Set DataTable("ACC_Pol_OffPhone","GL-Data")
		GL_Accident.WebEdit("ACC_Pol_NCIC").Set DataTable("ACC_Pol_NCIC","GL-Data")
	ElseIf ((DataTable("ACC_Fire","GL-Data") = "ON") OR (DataTable("ACC_Ambulance","GL-Data") = "ON") OR (DataTable("ACC_Other","GL-Data") = "ON")) Then
		GL_Accident.WebEdit("ACC_Ambu_AuthName").Set DataTable("ACC_Ambu_AuthName","GL-Data")
		GL_Accident.WebEdit("ACC_Ambu_Report").Set DataTable("ACC_Ambu_Report","GL-Data")
	End If
	GL_Accident.WebButton("Next>>").Click 
	Browser("ClaimsBrowser").Sync
	
End function


Function Verify_City_Length()
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>>  Site Details - Verify City Length"
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



Function Party()

	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Party Screen "
	Set GL_Party=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	
	If GL_Party.WebEdit("Party_Fname").Exist(10) Then
        
        For i = 1 to DataTable("No.Of.Claimants","GL-Data")
			GL_Party.WebEdit("Party_Fname").Set DataTable("Party_Fname","GL-Data")
			GL_Party.WebEdit("Party_MI").Set DataTable("Party_MI","GL-Data")
			GL_Party.WebEdit("Party_Lname").Set DataTable("Party_Lname","GL-Data")
			
			If  DataTable("IN_Category","GL-Data") <> "Inland Marine" Then
				If DataTable("IN_Category","GL-Data") <> "Innkeepers/Guest Property" Then
					GL_Party.WebCheckBox("Party_Injured").Set DataTable("Party_Injured","GL-Data")
				End If
			End If
			if DataTable("Party_Injured","GL-Data") = "ON" then
				GL_Party.WebCheckBox("Party_Fatality").Set DataTable("Party_Fatality","GL-Data")
			End if
			GL_Party.WebCheckBox("Party_PropertyDamage").Set DataTable("Party_PropertyDamage","GL-Data")
			GL_Party.WebCheckBox("Party_VehicleDamage").Set DataTable("Party_VehicleDamage","GL-Data")
			GL_Party.WebCheckBox("Party_Attorney").Set DataTable("Party_Attorney","GL-Data")
			GL_Party.WebButton("Party_Add_To_List").Click
			Browser("ClaimsBrowser").Sync
		Next
	End If
	GL_Party.WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	
End function

Function PartyInfo1()
	
	Set GL_Party=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> First Party - Party Info Screen "
	'Party  Info1
	Browser("ClaimsBrowser").Sync
	GL_Party.WebRadioGroup("PartyInfo1_ZurichEmp").Select DataTable("PartyInfo_ZurichEmp","GL-Data")  'value  true/false
	GL_Party.WebRadioGroup("PartyInfo1_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
	Browser("ClaimsBrowser").Sync
	If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "No" Then
    		GL_Party.WebEdit("PartyInfo1_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
			GL_Party.WebEdit("PartyInfo1_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
			GL_Party.WebEdit("PartyInfo1_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
			Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*NumberOfDependents","Index:=0").Set DataTable("Dependent_Count","GL-Data")
	End If
	GL_Party.WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> First Party - Injury Screen "
	'Injury  Info1
	If DataTable("Party_Injured","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
			GL_Party.WebEdit("Inj_Description1").Set DataTable("Inj_Description","GL-Data")
			GL_Party.WebEdit("Inj_CauseInjury1").Set DataTable("Inj_CauseInjury","GL-Data")
			If DataTable("Party_Fatality","GL-Data") = "ON" Then
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_DateOfDeath").Set DATE()
			End If
			GL_Party.WebList("Inj_Nature1").Select DataTable("Inj_Nature","GL-Data")
			GL_Party.WebList("Inj_BodyPart1").Select  DataTable("Inj_BodyPart","GL-Data")
			GL_Party.WebList("Inj_InitialTreatment1").Select DataTable("Inj_InitialTreatment","GL-Data")
			If DataTable("Inj_InitialTreatment","GL-Data") = "NO MEDICAL TREATMENT"  or DataTable("Inj_InitialTreatment","GL-Data") = "MINOR ON-SITE REMEDIES BY EMPLOYER MEDICAL STAFF" Then
					'do nothing
			else
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Physician.*LastName").Set "LN"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Physician.*FirstName").Set "FN"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Physician.*MiddleName").Set "M"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Physician.*addressLines.*","Index:=0").Set "Phy Addr1"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Physician.*addressLines.*","Index:=1").Set "Phy Addr2"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Physician.*postalCode").Set "12345"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Physician.*Fax").Set "111-111-1111"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Physician.*Phone").Set "222-222-2222"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Physician.*EmailAddress").Set "Phy@csc.com"
				
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Hospital.*FullName").Set "Hosp Name"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Hospital.*addressLines.*","Index:=0").Set "Addr1"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Hospital.*addressLines.*","Index:=1").Set "Addr2"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Hospital.*postalCode").Set "12345"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Hospital.*Phone").Set "111-111-1111"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Hospital.*Fax").Set "222-222-2222"
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("name:=.*Hospital.*EmailAddress").Set "Hosp@csc.com"
			End if	
			
			GL_Party.WebButton("Next>>").Click
	End If
	'Property Damage1
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> First Party - PD Screen "
	Browser("ClaimsBrowser").Sync
	If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
        	GL_Party.WebRadioGroup("PropertyDam1_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
        	
        	If DataTable("PropertyDam_Location","GL-Data") = "O" Then
	    		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=addressLines","Index:=2").Set "Address1"
	    		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=addressLines","Index:=3").Set "Address2"
	    		Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("html id:=postalCode","Index:=1").Set "12345"
    		End If
    		
			GL_Party.WebEdit("PropertyDam1_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
			GL_Party.WebEdit("PropertyDam1_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
			GL_Party.WebEdit("PropertyDam1_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
			GL_Party.WebList("PropertyDam1_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
			GL_Party.WebCheckBox("PropertyDam1_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
			GL_Party.WebCheckBox("PropertyDam1_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
			GL_Party.WebButton("Next>>").Click
			Browser("ClaimsBrowser").Sync
	End If
	
	If DataTable("Party_Attorney","GL-Data") = "ON" Then
		call Attorney()
		Browser("ClaimsBrowser").Sync
	End If 
   
   If DataTable("DriverData","GL-Data") = "Yes" Then
		call DriverData()
		call VehicleData()
		call VehicleDamage()
		call LossEvaluation()			
	End If 
	
End Function

Function PartyInfo2()
	
	Set GL_Party=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Second Party - Party Info Screen "
	'Party Info 2
		GL_Party.WebRadioGroup("PartyInfo2_ZurichEmp").Select DataTable("PartyInfo_ZurichEmp","GL-Data")  'value  true/false
		GL_Party.WebRadioGroup("PartyInfo2_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
		Browser("ClaimsBrowser").Sync
	If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
			GL_Party.WebEdit("PartyInfo2_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
			GL_Party.WebEdit("PartyInfo2_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
			GL_Party.WebEdit("PartyInfo2_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
	End if
	GL_Party.WebButton("Next>>").Click
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Second Party - Injury Screen "
	'Injury 2
	If DataTable("Party_Injured","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
			GL_Party.WebEdit("Inj_Description2").Set DataTable("Inj_Description","GL-Data")
			GL_Party.WebEdit("Inj_CauseInjury2").Set DataTable("Inj_CauseInjury","GL-Data")
			If DataTable("Party_Fatality","GL-Data") = "ON" Then
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_DateOfDeath").Set DATE()
			End If
			GL_Party.WebList("Inj_Nature2").Select DataTable("Inj_Nature","GL-Data")
			GL_Party.WebList("Inj_BodyPart2").Select  DataTable("Inj_BodyPart","GL-Data")
			GL_Party.WebList("Inj_InitialTreatment2").Select DataTable("Inj_InitialTreatment","GL-Data")
			GL_Party.WebButton("Next>>").Click
	End  if
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Second Party - PD Screen "
	'Property Damage 2
	If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
			GL_Party.WebRadioGroup("PropertyDam2_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
			GL_Party.WebEdit("PropertyDam2_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
			GL_Party.WebEdit("PropertyDam2_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
			GL_Party.WebEdit("PropertyDam2_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
			GL_Party.WebList("PropertyDam2_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
			GL_Party.WebCheckBox("PropertyDam2_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
			GL_Party.WebCheckBox("PropertyDam2_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
			GL_Party.WebButton("Next>>").Click
			Browser("ClaimsBrowser").Sync
	End if
	
	If DataTable("Party_Attorney","GL-Data") = "ON" Then
		call Attorney()
		Browser("ClaimsBrowser").Sync
	End If 
   
   If DataTable("DriverData","GL-Data") = "Yes" Then
		call DriverData()
		call VehicleData()
		call VehicleDamage()
		call LossEvaluation()			
	End If 
	
	
End Function

Function PartyInfo3()

	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Third Party - Party Info Screen "
	Set GL_Party=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	'Party Info 3
		GL_Party.WebRadioGroup("PartyInfo3_ZurichEmp").Select DataTable("PartyInfo_ZurichEmp","GL-Data")  'value  true/false
		GL_Party.WebRadioGroup("PartyInfo3_PartyAddSame_AccAdd").Select DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data")   ' value Yes/No
		Browser("ClaimsBrowser").Sync
	If DataTable("PartyInfo_PartyAddSame_AccAdd","GL-Data") = "NO" Then
		GL_Party.WebEdit("PartyInfo3_Add1").Set DataTable("PartyInfo_Add1","GL-Data")
		GL_Party.WebEdit("PartyInfo3_Add2").Set DataTable("PartyInfo_Add2","GL-Data")
		GL_Party.WebEdit("PartyInfo3_Zip").Set DataTable("PartyInfo_Zip","GL-Data")
	End if
		GL_Party.WebButton("Next>>").Click
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Third Party - Injury Screen "	
	'Injury 3
	If DataTable("Party_Injured","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
			GL_Party.WebEdit("Inj_Description3").Set DataTable("Inj_Description","GL-Data")
			GL_Party.WebEdit("Inj_CauseInjury3").Set DataTable("Inj_CauseInjury","GL-Data")
			If DataTable("Party_Fatality","GL-Data") = "ON" Then
				Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION").WebEdit("Inj_DateOfDeath").Set DATE()
			End If
			GL_Party.WebList("Inj_Nature3").Select DataTable("Inj_Nature","GL-Data")
			GL_Party.WebList("Inj_BodyPart3").Select  DataTable("Inj_BodyPart","GL-Data")
			GL_Party.WebList("Inj_InitialTreatment3").Select DataTable("Inj_InitialTreatment","GL-Data")
			GL_Party.WebButton("Next>>").Click
	End if
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Third Party - PD Screen "
	'Property Damage 3
	
	If DataTable("Party_PropertyDamage","GL-Data") = "ON" Then
			Browser("ClaimsBrowser").Sync
			GL_Party.WebRadioGroup("PropertyDam3_Location").Select DataTable("PropertyDam_Location","GL-Data")  'values  C,A,O
			GL_Party.WebEdit("PropertyDam3_PropDescription").Set DataTable("PropertyDam_PropDescription","GL-Data")
			GL_Party.WebEdit("PropertyDam3_DamDescription").set DataTable("PropertyDam_DamDescription","GL-Data")
			GL_Party.WebEdit("PropertyDam3_LossAmount").Set DataTable("PropertyDam_LossAmount","GL-Data")
			GL_Party.WebList("PropertyDam3_InsuranceInfo").Select DataTable("PropertyDam_InsuranceInfo","GL-Data")
			GL_Party.WebCheckBox("PropertyDam3_BuisnessInterption").Set DataTable("PropertyDam_BuisnessInterption","GL-Data")
			GL_Party.WebCheckBox("PropertyDam3_ExceedsTheshold_Amount").Set DataTable("PropertyDam_ExceedsTheshold_Amount","GL-Data")
			GL_Party.WebButton("Next>>").Click
			Browser("ClaimsBrowser").Sync
	End if 
	
End function 


Function  DriverData()

	Browser("ClaimsBrowser").Sync
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> 1st Party - Driver Data Screen "
	Set GL_DriData=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	Browser("ClaimsBrowser").Sync	
	GL_DriData.WebCheckBox("Driver_SameAs").Set DataTable("DriverData_SameAsOwner","GL-Data")
	Wait(2)
	If DataTable("DriverData_SameAsOwner","GL-Data")="OFF"  Then 
		GL_DriData.WebEdit("DriData_Fname").Set DataTable("DriverData_FName","GL-Data")
		GL_DriData.WebEdit("DriData_MI").Set DataTable("DriverData_MiddleName","GL-Data")
		GL_DriData.WebEdit("DriData_Lname").Set DataTable("DriverData_LastName","GL-Data")
		GL_DriData.WebEdit("DriData_Add1").Set DataTable("DriverData_Address1","GL-Data")
		GL_DriData.WebEdit("DriData_Add2").Set DataTable("DriverData_Address2","GL-Data")
	End If 
	GL_DriData.WebEdit("DriData_Zip").Set DataTable("DriverData_ZIP","GL-Data")
	GL_DriData.WebEdit("DriData_WorkPhone").Set DataTable("DriverData_Workphone","GL-Data")
	GL_DriData.WebEdit("DriData_Cell").Set DataTable("DriverData_Cellphone","GL-Data")
	GL_DriData.WebEdit("DriData_Fax").Set DataTable("DriverData_Fax","GL-Data")
	GL_DriData.WebEdit("DriData_Email").Set DataTable("DriverData_Email","GL-Data")
	GL_DriData.WebList("DriData_Distribution").Select DataTable("DriverData_DistributionPrefer","GL-Data")
	GL_DriData.WebEdit("DriData_SSN").Set DataTable("DriverData_SSN","GL-Data")
	GL_DriData.WebEdit("DriData_DOB").Set DataTable("DriverData_DOB","GL-Data")
	GL_DriData.WebList("DriData_Gender").Select DataTable("DriverData_Gender","GL-Data")
	GL_DriData.WebList("DriData_Marital").Select DataTable("DriverData_MaritalStatus","GL-Data")
	GL_DriData.WebEdit("DriData_Dependent").Set DataTable("DriverData_DependentCount","GL-Data")
	GL_DriData.WebEdit("DriData_Licence").Set DataTable("DriverData_DriverLicense","GL-Data")
	GL_DriData.WebList("DriData_StateofIssue").Select DataTable("DriverData_StateOfIssue","GL-Data")
    GL_DriData.WebButton("Next>>").Click
    
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

End Function

Function  VehicleData()
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> 1st Party - Vehicle Data Screen "
	Set GL_Frame=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") '
	GL_Frame.WebEdit("VehicleData_VIN").Set DataTable("VehicleData_VIN","GL-Data")
	Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Attorney").Image("VIN_Image").Click
	GL_Frame.WebButton("Next>>").Click
End Function

Function VehicleDamage()
	
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> 1st Party - Vehicle Damage Screen "
	
	Set GL_VehDamage=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	GL_VehDamage.WebList("VehicleDamage_EstimateSpeed").Select DataTable("VehicleDamage_EstimateSpeed","GL-Data")
	GL_VehDamage.WebList("VehicleDamage_LossType").Select DataTable("VehicleDamage_LossType","GL-Data")
	GL_VehDamage.WebRadioGroup("VDamage_Area").Select DataTable("VDamage_Area","GL-Data")
	GL_VehDamage.WebList("VehicleDamage_PersonalProperty").Select DataTable("VehicleDamage_PersonalProperty","GL-Data")
	GL_VehDamage.WebButton("Next>>").Click
	
End Function

Function LossEvaluation()

	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Loss Evaluation Screen "
	Set GL_Frame=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 	
	GL_Frame.WebButton("Next>>").Click
	
End Function


Function Witness()

	Set GL_Witness=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION")
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Witness Screen "	
	GL_Witness.WebList("WitnessList").Select DataTable("Witness_List","GL-Data")
	If DataTable("Witness_List","GL-Data") = "Yes" Then	
		GL_Witness.WebEdit("Wit_FirstName").Set DataTable("Witness_FirstName","GL-Data")
		GL_Witness.WebEdit("Wit_LastName").Set DataTable("Witness_LastName","GL-Data")
		GL_Witness.WebEdit("Wit_Zip").Set DataTable("Witness_Zip","GL-Data")
	End If
	
	GL_Witness.WebButton("Next>>").Click
	
End Function

Function Dependent()
	
	Set GL_Party=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	Environment.value("str_ScreenName") = "Carepoint - GL >>>> First Party - Dependent Screen "
	
	GL_Party.WebEdit("name:=.*Dependant.*FirstName","Index:=0").Set "Dependent FN"
	GL_Party.WebEdit("name:=.*Dependant.*LastName","Index:=0").Set "Dependent LN"
	GL_Party.WebEdit("name:=.*Dependant.*MI","Index:=0").Set "M"
	GL_Party.WebEdit("name:=.*Dependant.*DOB","Index:=0").Set "10/10/1979"
	GL_Party.WebEdit("name:=.*Dependant.*addressLines.*","Index:=0").Set "Addr1"
	GL_Party.WebEdit("name:=.*Dependant.*addressLines.*","Index:=1").Set "Addr2"
	GL_Party.WebEdit("name:=.*Dependant.*postalCode","Index:=0").Set "12345"
	GL_Party.WebEdit("name:=.*Dependant.*Phone","Index:=0").Set "111-111-1111"
	GL_Party.WebList("name:=.*Dependant.*RelationCode","Index:=0").Select "#1"	
	GL_Party.WebButton("Next>>").Click
	
End Function


Function Attorney()

	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Attorney Screen "
	Set GL_Frame=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	Set GL_Attorney=Browser("ClaimsBrowser").Page("CCC Manager Portal 7.1").Frame("Attorney") 
	GL_Attorney.WebList("AttorneyList").Select DataTable("Attorney_List","GL-Data")
	If DataTable("Attorney_List","GL-Data") = "Yes" Then
		GL_Attorney.WebEdit("Attorney_FirmName").Set DataTable("Attorney_FirmName","GL-Data")
		GL_Attorney.WebEdit("Attorney_FirstName").Set DataTable("Attorney_FirstName","GL-Data")
		GL_Attorney.WebEdit("Attorney_LastName").Set DataTable("Attorney_LastName","GL-Data")
		GL_Attorney.WebEdit("Attorney_Address1").Set DataTable("Attorney_Address1","GL-Data")
		GL_Attorney.WebEdit("Attorney_Address2").Set "Address2"
		GL_Attorney.WebEdit("Attorney_Phone").Set "123-343-4343"
		GL_Attorney.WebEdit("Attorney_ZIP").Set DataTable("Attorney_ZIP","GL-Data")
		GL_Attorney.WebEdit("Attorney_Email").Set DataTable("Attorney_Email","GL-Data")
		GL_Attorney.WebEdit("Attorney_Fax").Set DataTable("Attorney_Fax","GL-Data")
	End If
	GL_Frame.WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	
End Function


Function Additional_Information()

	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Additional Info Screen "
	Set GL_AddInfo=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	GL_AddInfo.WebButton("Next>>").Click
	Browser("ClaimsBrowser").Sync
	
End Function

Function Assignment()

	Environment.value("str_ScreenName") = "Carepoint - GL >>>> Assignment Screen "

	Set GL_Assignmen=Browser("ClaimsBrowser").Page("Inbox").Frame("DIACTION") 
	If GL_Assignmen.WebButton("Run Assignment").Exist(5) Then 
		GL_Assignmen.WebList("ACC_AccCode").Select DataTable("ACC_AccCode_Assignment","GL-Data")		
		GL_Assignmen.WebButton("Run Assignment").Click
		Browser("ClaimsBrowser").sync
	End If
	GL_Assignmen.WebButton("Ass_Save").Click
	Browser("ClaimsBrowser").sync	
	GL_Assignmen.WebButton("Get_Claim_Number").Click
	Browser("ClaimsBrowser").sync	
	
	If Not Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html id:=PegaGadget0Ifr","html tag:=IFRAME").WebButton("class:=buttonTdButton","html tag:=BUTTON","name:=No Duplicates Found").Exist(5)  Then
		If GL_Assignmen.WebList("Reason").Exist(5) Then
			If GL_Assignmen.WebButton("Override").GetROProperty("width")>0 Then
				GL_Assignmen.WebEdit("Contact_Name").Set DataTable("Party_Fname","GL-Data")
				GL_Assignmen.WebEdit("Contact_Phone").Set DataTable("PartyInfo_Fax","GL-Data")
				GL_Assignmen.WebList("Reason").Select "Other"
				GL_Assignmen.WebEdit("Reason_Edit").Set "other"
				GL_Assignmen.WebButton("Override").Click
				Browser("ClaimsBrowser").Sync
			End If	
				
		End If
		If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html id:=PegaGadget0Ifr","html tag:=IFRAME").WebButton("class:=buttonTdButton","html tag:=BUTTON","name:=No Duplicates Found").Exist(5) Then
			Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html id:=PegaGadget0Ifr","html tag:=IFRAME").WebButton("class:=buttonTdButton","html tag:=BUTTON","name:=No Duplicates Found").Click
		End If
		If Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Exist(5) then
				Browser("ClaimsBrowser").Dialog("Message from webpage").WinButton("OK").Click				
		End if
	Else
		If Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html id:=PegaGadget0Ifr","html tag:=IFRAME").WebButton("class:=buttonTdButton","html tag:=BUTTON","name:=No Duplicates Found").Exist(5) Then
			Browser("name:=CCC.*").Page("title:=CCC.*").Frame("html id:=PegaGadget0Ifr","html tag:=IFRAME").WebButton("class:=buttonTdButton","html tag:=BUTTON","name:=No Duplicates Found").Click
		End If
	End If	
	Browser("ClaimsBrowser").sync	
	
End Function

Function GetClaimNumber()

	Claim_Number= Browser("ClaimsBrowser").Page("Inbox").Frame("Review_Distribution_Frame").WebTable("ClaimNumber_Table").GetROProperty("innertext")
	Claim_Number=Trim(Claim_Number)
	Claim_Number=right(Claim_Number,10)
	Environment.Value("Claim_Number") = Claim_Number
	Environment.Value("NewClaimNumber") =  Claim_Number & "   " & Environment.Value("SCaseId")
	Print "+++++++++++++++++++++++++++++++++++++ Claim Number is +++++++++++++++ " & Environment.Value("NewClaimNumber")  & " ++++++++++++++++++++++++++++++++++++++++++++++++++++++++ "
	
End function

Function Logout()
	
	Environment.value("str_ScreenName") = "Carepoint - GL  >>>> Logoff Screen "
	
	Browser("name:=CC.*").Page("title:=CC.*").Image("name:=Image","image type:=Image Link","Index:=0").Click
	Browser("name:=CC.*").Page("title:=CC.*").WebElement("innertext:=Log off","html id:=ItemMiddle").Click
	SystemUtil.CloseProcessByName "iexplore.exe"
	
End Function


'Created By :-  Srirekha Talasila
'This will handle Distributions in Review Screen 

Function Review_Distribution()
	
		Environment.value("str_ScreenName") = "Carepoint - GL  >>>> Review Distribution Screen "
		On Error Resume Next
		
		Browser("name:=CCC.*").Sync
		Call GetClaimNumber()
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

'Created By :-  Srirekha Talasila
'This will Verify Search Functionality using S-Case and Claim Number
Function Binocular_Search()

	Environment.value("str_ScreenName") = "Carepoint - GL  >>>> Binocular Search Screen "
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



