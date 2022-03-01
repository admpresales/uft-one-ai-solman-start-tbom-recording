Dim oShell, counter

BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3												'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")												'Set the variable for what application (in this case the browser) we are acting upon

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("WorkItemsURL")												'Navigate to the application URL
AppContext.Maximize																		'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																			'Wait for the browser to stop spinning
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application
If AIUtil("text_box", "User").Exist Then
	AIUtil("text_box", "User").Type DataTable.Value("Login")
	AIUtil("text_box", "Password").Type DataTable.Value("Password")
	AIUtil("button", "Log On").Click
	counter = 0
	Do
		counter = counter + 1
		wait 1
		If counter >= 60 Then
			'msgbox "The search icon didn't show up within " & counter & " tries, check application."
			Reporter.ReportEvent micFail, "Login", "Login failed.  The search icon didn't show up within " & counter & " tries, check application."
			ExitTestIteration
		End If
		If counter>=3 Then
			If AIUtil("button", "Continue").Exist(0) Then
				'AIUtil("check_box", "Cancel all existing Iogons").SetState "Off"
				AIUtil("button", "Continue").Click
			End If
		End If
	Loop Until AIUtil("search").Exist(0)
End If

AIUtil("gear_settings").Click
AIUtil.FindTextBlock("Filter").Click
AIUtil("combobox", "Reset").Type "ID"
AIUtil("text_box", micAnyText, micWithAnchorOnLeft, AIUtil("combobox", "Reset")).Click
AIUtil("text_box", micAnyText, micWithAnchorOnLeft, AIUtil("combobox", "Reset")).Type DataTable.Value("WorkItemNumber")
AIUtil("button", "OK").Click
counter = 0
Do
	counter = counter + 1
	AIUtil("check_box", micAnyText, micFromTop, 2).SetState "On"
	If counter >= 60 Then
		'msgbox "The search icon didn't show up within " & counter & " tries, check application."
		Reporter.ReportEvent micFail, "Click WorkItemNumber", "The WorkItemNumber link "&  DataTable.Value("WorkItemNumber") & " didn't show up within " & counter & " tries, check application."
		ExitTestIteration
	End If
Loop Until AIUtil.FindText(DataTable.Value("WorkItemNumber")).Exist(0)

counter = 0
Do
	counter = counter + 1
	AIUtil.FindText(DataTable.Value("WorkItemNumber")).Click	
	If counter >= 60 Then
		'msgbox "The search icon didn't show up within " & counter & " tries, check application."
		Reporter.ReportEvent micFail, "Click WorkItemNumber", "The WorkItemNumber link "&  DataTable.Value("WorkItemNumber") & " didn't show up within " & counter & " tries, check application."
		ExitTestIteration
	End If
Loop Until AIUtil("button", "Edit").Exist(0)

AIUtil("button", "Edit").Click
AIUtil("button", "Create TBOM").Click
Set AppContext=Browser("CreationTime:=1")												'Set the variable for what application (in this case the browser) we are acting upon
AppContext.Sync																			'Wait for the browser to stop spinning

Set oShell = CreateObject ("WSCript.shell")
oShell.run "C:\Users\demo\Downloads\ags_work_appln.sap"
Set oShell = Nothing
counter = 0
Do
	counter = counter + 1
	wait 1
	If counter >= 60 Then
		'msgbox "The search icon didn't show up within " & counter & " tries, check application."
		Reporter.ReportEvent micFail, "TBOM Recording", "The TBOM Recording window didn't show up within " & counter & " tries, check application."
		ExitTestIteration
	End If
Loop Until SAPGuiSession("Session").SAPGuiWindow("TBOM Recording for Transaction").SAPGuiButton("Start Recording").Exist(0)
SAPGuiSession("Session").SAPGuiWindow("TBOM Recording for Transaction").SAPGuiButton("Start Recording").Click

